import os
import json
import base64
import logging
import re
import io
import asyncio
from datetime import datetime, timezone, timedelta
from telegram import Update, ReplyKeyboardMarkup, InputFile
from telegram.ext import (
    Application, CommandHandler, MessageHandler, filters,
    ContextTypes, ConversationHandler
)
import httpx
import gspread
import openpyxl

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
WEBHOOK_URL = os.environ.get("WEBHOOK_URL", "")
PORT = int(os.environ.get("PORT", 8080))

# --- Google Sheets ---
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
REPORT_SPREADSHEET_ID = os.environ.get("REPORT_SPREADSHEET_ID", "")  # файл "Реализация 2026г"
GOOGLE_CREDENTIALS = os.environ.get("GOOGLE_CREDENTIALS", "")  # содержимое JSON-ключа целиком

SHEET_PRODUCTION = "Производство"
SHEET_SALES = "Реализация"

# Список разрешённых Telegram ID (закрытый доступ).
# Пустой список = пускать всех (страховка, чтобы не заблокировать себя).
_allowed_raw = os.environ.get("ALLOWED_USERS", "")
ALLOWED_USERS = {int(x) for x in _allowed_raw.replace(" ", "").split(",") if x.strip().isdigit()}

# --- Автобэкап ---
BACKUP_CHANNEL_ID = os.environ.get("BACKUP_CHANNEL_ID", "")
KZ_TZ = timezone(timedelta(hours=5))
BACKUP_HOUR = 22
_last_backup_counts = {}

# --- настройки отчёта реализации ---------------------------------
REPORT_TEMPLATE = "ШАБЛОН"                 # лист-образец (пустой бланк)
REPORT_DATE_COL = 2                        # B = Дата
REPORT_BUYER_COL = 3                       # C = Контрагент
# Для каждой фракции: (Кол-во, Цена, Сумма) — номера столбцов (1-based)
REPORT_FRAC_COLS = {
    "0-1": (4, 5, 6),     # D E F
    "1-2": (7, 8, 9),     # G H I
    "2-4": (10, 11, 12),  # J K L
    "4-6": (13, 14, 15),  # M N O
    "6-8": (16, 17, 18),  # P Q R
}
# Блоки (строки данных, без строки "Итого")
REPORT_BLOCKS = {
    "Безналичный": (8, 23),
    "Наличный": (26, 40),
}
RU_MONTHS = {
    1: "январь", 2: "февраль", 3: "март", 4: "апрель", 5: "май", 6: "июнь",
    7: "июль", 8: "август", 9: "сентябрь", 10: "октябрь", 11: "ноябрь", 12: "декабрь",
}
# ------------------------------------------------------------------

# Состояния диалогов ручного ввода
(P_DATE, P_FIO, P_TIRES, P_BAGS, P_THREAD, P_F01, P_F12, P_F24, P_F46, P_F68, P_CORD, P_NOTE) = range(12)
(S_DATE, S_BUYER, S_PAYTYPE, S_FRAC, S_KG, S_PRICE, S_SUM_VAT, S_VAT, S_NOTE) = range(12, 21)
# NEW: состояние подтверждения дубля при фото-производстве
# + доспрос нал/безнал после фото-реализации (чтобы продажа попала в отчёт)
(PHOTO_PROD_CONFIRM, PHOTO_SALE_PAY, PHOTO_TYPE_CONFIRM) = range(21, 24)
(E_DATE, E_AMOUNT, E_GROUP, E_CATEGORY, E_SOURCE, E_CONTRAGENT, E_NOTE, E_CONFIRM) = range(24, 32)

FRACTIONS = ["0-1", "1-2", "2-4", "4-6", "6-8"]

# Заголовки листов склада (порядок колонок = порядок записи)
PROD_HEADERS = ["Дата", "ФИО оператора", "Вес шин кг", "Мешки шт", "Нитки",
                "Фракция 0-1", "Фракция 1-2", "Фракция 2-4", "Фракция 4-6", "Фракция 6-8",
                "Всего крошки кг", "Металлокорд кг", "Примечание"]
SALES_HEADERS = ["Дата", "Покупатель", "Тип расчёта", "Фракция", "Количество кг",
                 "Цена за кг", "Сумма с НДС", "Сумма НДС", "Примечание"]

_gc = None
_spreadsheet = None
_report_spreadsheet = None


def parse_num(text):
    """Превращает '210', '210,5', ' 210 кг ' в число. Пустое/прочерк -> 0."""
    if text is None:
        return 0
    t = str(text).strip().lower().replace(",", ".")
    t = t.replace("кг", "").replace("тнг", "").replace(" ", "")
    if t in ("", "-", "—", "нет"):
        return 0
    try:
        num = float(re.sub(r"[^0-9.]", "", t))
        return int(num) if num == int(num) else num
    except Exception:
        return 0


def parse_date_or_none(date_text):
    """Главная проверка даты. Принимает строку, которую отдал Claude или ввёл человек.
    Понимает обычные форматы С РАЗДЕЛИТЕЛЯМИ: '08.06.2026', '8.6.26', '8/6/2026', '8-6-26'.
    Год из двух цифр -> 20xx. Возвращает datetime, либо None если разобрать нельзя.
    НИКАКОГО угадывания слитных цифр — если непонятно, вернём None и спросим человека."""
    if date_text is None:
        return None
    s = str(date_text).strip()
    if not s:
        return None
    # сперва пробуем строгие форматы целиком
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    # затем — разделители любого вида: ровно 3 числовые части день/месяц/год
    parts = [p for p in re.split(r"[.\-/\s]+", s) if p != ""]
    if len(parts) == 3 and all(p.isdigit() for p in parts):
        d, m, y = parts
        try:
            d_i, m_i, y_i = int(d), int(m), int(y)
            if y_i < 100:
                y_i += 2000
            return datetime(y_i, m_i, d_i)
        except Exception:
            return None
    return None


def date_to_str(dt):
    """datetime -> 'дд.мм.гггг'."""
    return dt.strftime("%d.%m.%Y")


# ---------- контроль доступа ----------

def is_allowed(update: Update) -> bool:
    """True, если пользователь в списке разрешённых.
    Если ALLOWED_USERS пуст — пускаем всех (страховка от самоблокировки)."""
    if not ALLOWED_USERS:
        return True
    user = update.effective_user
    return bool(user and user.id in ALLOWED_USERS)


# ---------- Google Sheets: подключение и доступ к листам ----------

def _get_client():
    global _gc
    if _gc is None:
        if not GOOGLE_CREDENTIALS:
            raise RuntimeError("Не задан GOOGLE_CREDENTIALS в переменных окружения.")
        creds_dict = json.loads(GOOGLE_CREDENTIALS)
        _gc = gspread.service_account_from_dict(creds_dict)
    return _gc


def get_spreadsheet():
    global _spreadsheet
    if _spreadsheet is not None:
        return _spreadsheet
    if not SPREADSHEET_ID:
        raise RuntimeError("Не задан SPREADSHEET_ID в переменных окружения.")
    _spreadsheet = _get_client().open_by_key(SPREADSHEET_ID)
    return _spreadsheet


def get_report_spreadsheet():
    global _report_spreadsheet
    if _report_spreadsheet is not None:
        return _report_spreadsheet
    if not REPORT_SPREADSHEET_ID:
        raise RuntimeError("Не задан REPORT_SPREADSHEET_ID в переменных окружения.")
    _report_spreadsheet = _get_client().open_by_key(REPORT_SPREADSHEET_ID)
    return _report_spreadsheet


def get_worksheet(title, headers):
    sh = get_spreadsheet()
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=1000, cols=max(len(headers), 12))
        ws.append_row(headers, value_input_option="USER_ENTERED")
        return ws
    first = ws.row_values(1)
    if not first:
        ws.append_row(headers, value_input_option="USER_ENTERED")
    return ws


def get_production_ws():
    return get_worksheet(SHEET_PRODUCTION, PROD_HEADERS)


def get_sales_ws():
    return get_worksheet(SHEET_SALES, SALES_HEADERS)


# ---------- Чтение/запись данных ----------

def _fractions_signature(data):
    """NEW: подпись записи по фракциям (для сравнения дублей).
    Кортеж из 5 чисел фракций — устойчив к разнице в ФИО/примечании."""
    return (
        parse_num(data.get("фракция_0_1", 0)),
        parse_num(data.get("фракция_1_2", 0)),
        parse_num(data.get("фракция_2_4", 0)),
        parse_num(data.get("фракция_4_6", 0)),
        parse_num(data.get("фракция_6_8", 0)),
    )


def find_duplicate_production(data):
    """NEW: ищет в складе строку с той же датой И теми же фракциями.
    Возвращает True, если такая уже есть (вероятный дубль того же бланка).
    Разные смены за один день с разными цифрами дублем НЕ считаются."""
    ws = get_production_ws()
    rows = ws.get_all_values()
    if len(rows) <= 1:
        return False

    target_dt = parse_date_or_none(data.get("дата", ""))
    target_sig = _fractions_signature(data)

    for r in rows[1:]:
        row_date = r[0] if len(r) > 0 else ""
        # сравниваем даты по смыслу (а не по тексту): 08.06.2026 == 8.6.26
        row_dt = parse_date_or_none(row_date)
        same_date = False
        if target_dt and row_dt:
            same_date = (target_dt.date() == row_dt.date())
        else:
            same_date = (str(row_date).strip() == str(data.get("дата", "")).strip())
        if not same_date:
            continue
        # сравниваем фракции (индексы 5..9 в PROD_HEADERS)
        row_sig = tuple(
            parse_num(r[i]) if i < len(r) else 0
            for i in range(5, 10)
        )
        if row_sig == target_sig:
            return True
    return False


def save_production(data):
    ws = get_production_ws()
    f01 = parse_num(data.get("фракция_0_1", 0))
    f12 = parse_num(data.get("фракция_1_2", 0))
    f24 = parse_num(data.get("фракция_2_4", 0))
    f46 = parse_num(data.get("фракция_4_6", 0))
    f68 = parse_num(data.get("фракция_6_8", 0))
    total = f01 + f12 + f24 + f46 + f68
    bags = round(total / 30) if total else 0  # 1 мешок = 30 кг; считаем сами, не из распознавания
    row = [
        data.get("дата", datetime.now().strftime("%d.%m.%Y")),
        data.get("фио", ""),
        parse_num(data.get("вес_шин", 0)),
        bags,
        parse_num(data.get("нитки", 0)),
        f01, f12, f24, f46, f68,
        total,
        parse_num(data.get("металл_корд", 0)),
        data.get("примечание", ""),
    ]
    ws.append_row(row, value_input_option="USER_ENTERED")


def save_sale(data):
    ws = get_sales_ws()
    qty_kg = parse_num(data.get("количество_кг", 0))
    if not qty_kg and data.get("количество_т"):
        qty_kg = parse_num(data.get("количество_т", 0)) * 1000
    row = [
        data.get("дата", datetime.now().strftime("%d.%m.%Y")),
        data.get("покупатель", ""),
        data.get("тип_расчета", ""),
        data.get("фракция", ""),
        qty_kg,
        parse_num(data.get("цена_за_кг", 0)),
        parse_num(data.get("сумма_с_ндс", 0)),
        parse_num(data.get("сумма_ндс", 0)),
        data.get("примечание", ""),
    ]
    ws.append_row(row, value_input_option="USER_ENTERED")


# ---------- запись реализации в отчёт "Реализация 2026г" ----------

def _normalize_fraction(frac_text):
    for f in FRACTIONS:
        if f in str(frac_text):
            return f
    return None


def _get_or_create_month_sheet(sh, dt):
    title = f"{RU_MONTHS[dt.month]} {dt.year}"
    try:
        return sh.worksheet(title)
    except gspread.WorksheetNotFound:
        pass
    template = sh.worksheet(REPORT_TEMPLATE)
    new_ws = template.duplicate(new_sheet_name=title)
    return new_ws


def write_sale_to_report(data):
    """Дописывает строку реализации в отчёт. Возвращает (ok: bool, msg: str)."""
    pay = str(data.get("тип_расчета", "")).strip()
    if pay not in REPORT_BLOCKS:
        return False, f"тип расчёта не распознан ('{pay}') — строка в отчёт не добавлена"

    frac = _normalize_fraction(data.get("фракция", ""))
    if frac is None:
        return False, "фракция не распознана — строка в отчёт не добавлена"

    dt = parse_date_or_none(data.get("дата", ""))
    if dt is None:
        return False, "дата не распознана — строка в отчёт не добавлена"

    sh = get_report_spreadsheet()
    ws = _get_or_create_month_sheet(sh, dt)

    start_row, end_row = REPORT_BLOCKS[pay]
    date_col_vals = ws.col_values(REPORT_DATE_COL)
    target_row = None
    for r in range(start_row, end_row + 1):
        val = date_col_vals[r - 1] if r - 1 < len(date_col_vals) else ""
        if not str(val).strip():
            target_row = r
            break
    if target_row is None:
        return False, f"блок «{pay}» заполнен (нет свободных строк) — строка не добавлена"

    qty_kg = parse_num(data.get("количество_кг", 0))
    if not qty_kg and data.get("количество_т"):
        qty_kg = parse_num(data.get("количество_т", 0)) * 1000
    price = parse_num(data.get("цена_за_кг", 0))
    total_sum = parse_num(data.get("сумма_с_ндс", 0))

    qcol, pcol, scol = REPORT_FRAC_COLS[frac]

    updates = [
        {"range": gspread.utils.rowcol_to_a1(target_row, REPORT_DATE_COL),
         "values": [[dt.strftime("%d.%m.%Y")]]},
        {"range": gspread.utils.rowcol_to_a1(target_row, REPORT_BUYER_COL),
         "values": [[data.get("покупатель", "")]]},
        {"range": gspread.utils.rowcol_to_a1(target_row, qcol), "values": [[qty_kg]]},
        {"range": gspread.utils.rowcol_to_a1(target_row, pcol), "values": [[price]]},
        {"range": gspread.utils.rowcol_to_a1(target_row, scol), "values": [[total_sum]]},
    ]
    ws.batch_update(updates, value_input_option="USER_ENTERED")
    return True, f"добавлено в «{ws.title}», блок «{pay}», строка {target_row}"


# ---------- Остаток ----------

def calc_stock():
    income = {k: 0 for k in FRACTIONS}
    ws_p = get_production_ws()
    rows = ws_p.get_all_values()
    for r in rows[1:]:
        for i, key in enumerate(FRACTIONS, 5):
            if i < len(r):
                try:
                    income[key] += float(str(r[i]).replace(",", ".")) if r[i] else 0
                except Exception:
                    pass

    outcome = {k: 0 for k in FRACTIONS}
    ws_s = get_sales_ws()
    rows_s = ws_s.get_all_values()
    for r in rows_s[1:]:
        frac = r[3] if len(r) > 3 else ""
        qty = r[4] if len(r) > 4 else 0
        for key in outcome:
            if key in str(frac):
                try:
                    outcome[key] += float(str(qty).replace(",", ".")) if qty else 0
                except Exception:
                    pass
                break

    stock = {k: income[k] - outcome[k] for k in income}
    return income, outcome, stock


# ---------- Распознавание фото через Claude ----------

async def call_claude(image_b64: str, prompt: str):
    async with httpx.AsyncClient(timeout=60) as client:
        response = await client.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json"
            },
            json={
                "model": "claude-opus-4-5",
                "max_tokens": 2000,
                "messages": [{
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": image_b64}},
                        {"type": "text", "text": prompt}
                    ]
                }]
            }
        )
        result = response.json()
        logger.info(f"Claude response: {result}")
        text = result["content"][0]["text"].strip()
        text = re.sub(r'```json|```', '', text).strip()
        return json.loads(text)


async def recognize_production(image_bytes: bytes):
    image_b64 = base64.b64encode(image_bytes).decode()
    prompt = """Это фото ежедневного отчёта по производству резиновой крошки.
На фото может быть одна или несколько строк. Верни ТОЛЬКО JSON без markdown.
Если одна строка — объект, если несколько — массив объектов. Формат:
{"дата":"дд.мм.гггг","фио":"ФИО","вес_шин":0,"мешки":0,"нитки":0,"фракция_0_1":0,"фракция_1_2":0,"фракция_2_4":0,"фракция_4_6":0,"фракция_6_8":0,"всего_итог":0,"металл_корд":0,"примечание":""}

ВАЖНО про ДАТУ:
- На бланке дату часто пишут слитно, без нулей и точек: формат деньмесяцгод.
  Пример: "5626" означает 5.6.26 -> верни "05.06.2026".
  Пример: "151226" означает 15.12.26 -> верни "15.12.2026".
- Год пишут двумя цифрами (26), это 2026 -> в ответе год всегда 4 цифры (2026).
- В ответе дата ВСЕГДА в виде "дд.мм.гггг" с точками и ведущими нулями.

ВАЖНО про ФРАКЦИИ:
- В одной ячейке фракции может стоять ДВА числа: верхнее = мешки (штуки), нижнее = килограммы.
- Бери ВСЕГДА НИЖНЕЕ число (килограммы). Верхнее (мешки) игнорируй.
- Если число одно — это килограммы.

ВАЖНО про ВСЕГО:
- На бланке есть итоговое поле «всего крошки» (общий вес за смену в кг).
- Прочитай это число и верни в "всего_итог" (если два числа — бери нижнее, кг). Если поля нет — 0.

Если что-то не читается — ставь 0."""
    return await call_claude(image_b64, prompt)


async def recognize_sale(image_bytes: bytes):
    image_b64 = base64.b64encode(image_bytes).decode()
    prompt = """Это фото накладной на отпуск/реализацию резиновой крошки.
Распознай данные и верни ТОЛЬКО JSON без markdown:
{"дата":"дд.мм.гггг","покупатель":"название организации или ИП","фракция":"например 2-4","количество_т":0.0,"количество_кг":0,"цена_за_кг":0,"сумма_с_ндс":0,"сумма_ндс":0,"примечание":""}
Количество бери из колонки "отпущено". Если оно в тоннах — заполни количество_т, если в кг — количество_кг.
Цену бери из колонки "цена" (цена за 1 кг). Бери ровно как в накладной, не пересчитывай.
Сумму с НДС бери из колонки "Сумма с НДС" (полная сумма включая НДС).
Сумму НДС бери из колонки "Сумма НДС" (только НДС). Если НДС в накладной нет — поставь 0.
Дату верни в виде "дд.мм.гггг". Если не читается — 0."""
    return await call_claude(image_b64, prompt)


# ---------- Команды и меню ----------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("⛔ Нет доступа. Обратитесь к администратору.")
        return
    keyboard = [
        ["📸 Производство", "📄 Реализация"],
        ["✍️ Ввод производства", "✍️ Ввод реализации"],
        ["💸 Ввод расхода"],
    ]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text(
        "👋 Привет! Я бот учёта резиновой крошки.\n\n"
        "📸 Производство — распознать фото отчёта\n"
        "📄 Реализация — распознать фото накладной\n"
        "✍️ Ввод производства — внести вручную по шагам\n"
        "✍️ Ввод реализации — внести вручную по шагам\n\n"
        "Команды:\n"
        "/ostatok — остаток на складе\n"
        "/last — последние записи\n"
        "/cancel — отменить ручной ввод",
        reply_markup=reply_markup
    )


async def ostatok(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("⛔ Нет доступа. Обратитесь к администратору.")
        return
    income, outcome, stock = calc_stock()
    total_in = sum(income.values())
    total_out = sum(outcome.values())
    total_stock = sum(stock.values())

    text = "📦 *ОСТАТОК НА СКЛАДЕ*\n\n"
    text += "```\n"
    text += f"{'Фракция':<8} {'Приход':>8} {'Расход':>8} {'Остаток':>9}\n"
    text += "-" * 37 + "\n"
    for key in FRACTIONS:
        text += f"{key:<8} {income[key]:>8.0f} {outcome[key]:>8.0f} {stock[key]:>9.0f}\n"
    text += "-" * 37 + "\n"
    text += f"{'ИТОГО':<8} {total_in:>8.0f} {total_out:>8.0f} {total_stock:>9.0f}\n"
    text += "```\n_Все данные в кг_"

    await update.message.reply_text(text, parse_mode="Markdown")


async def last_records(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("⛔ Нет доступа. Обратитесь к администратору.")
        return
    ws = get_production_ws()
    rows = ws.get_all_values()
    data_rows = rows[1:]
    if not data_rows:
        await update.message.reply_text("📭 Записей пока нет.")
        return
    last5 = data_rows[-5:]
    text = "📋 *Последние записи производства:*\n\n"
    for r in last5:
        date = r[0] if len(r) > 0 and r[0] else "—"
        fio = r[1] if len(r) > 1 and r[1] else "—"
        total = r[10] if len(r) > 10 and r[10] else "—"
        text += f"📅 {date} | {fio} | {total} кг\n"
    await update.message.reply_text(text, parse_mode="Markdown")


# ---------- Ручной ввод: ПРОИЗВОДСТВО ----------

async def manual_prod_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("⛔ Нет доступа. Обратитесь к администратору.")
        return ConversationHandler.END
    context.user_data["prod"] = {}
    await update.message.reply_text(
        "✍️ Ручной ввод производства.\n\n"
        "Введи *дату* (дд.мм.гггг).\n"
        "Или напиши «сегодня».\n\n"
        "Для отмены — /cancel",
        parse_mode="Markdown"
    )
    return P_DATE


async def manual_prod_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    t = update.message.text.strip()
    if t.lower() == "сегодня":
        context.user_data["prod"]["дата"] = datetime.now().strftime("%d.%m.%Y")
    else:
        dt = parse_date_or_none(t)
        if dt is None:
            await update.message.reply_text(
                "⚠️ Не понял дату. Введи в виде ДД.ММ.ГГГГ (например 05.06.2026)\n"
                "или напиши «сегодня». Для отмены — /cancel"
            )
            return P_DATE
        context.user_data["prod"]["дата"] = date_to_str(dt)
    await update.message.reply_text("👤 ФИО оператора?")
    return P_FIO


async def manual_prod_fio(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["prod"]["фио"] = update.message.text.strip()
    await update.message.reply_text("⚖️ Вес шин, кг?")
    return P_TIRES


async def manual_prod_tires(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["prod"]["вес_шин"] = parse_num(update.message.text)
    await update.message.reply_text("🧵 Нитки, кг? (или «-» если нет)")
    return P_THREAD


async def manual_prod_thread(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["prod"]["нитки"] = parse_num(update.message.text)
    await update.message.reply_text("Фракция 0-1, кг?")
    return P_F01


async def manual_prod_f01(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["prod"]["фракция_0_1"] = parse_num(update.message.text)
    await update.message.reply_text("Фракция 1-2, кг?")
    return P_F12


async def manual_prod_f12(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["prod"]["фракция_1_2"] = parse_num(update.message.text)
    await update.message.reply_text("Фракция 2-4, кг?")
    return P_F24


async def manual_prod_f24(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["prod"]["фракция_2_4"] = parse_num(update.message.text)
    await update.message.reply_text("Фракция 4-6, кг?")
    return P_F46


async def manual_prod_f46(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["prod"]["фракция_4_6"] = parse_num(update.message.text)
    await update.message.reply_text("Фракция 6-8, кг?")
    return P_F68


async def manual_prod_f68(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["prod"]["фракция_6_8"] = parse_num(update.message.text)
    await update.message.reply_text("🔩 Металлокорд, кг?")
    return P_CORD


async def manual_prod_cord(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["prod"]["металл_корд"] = parse_num(update.message.text)
    await update.message.reply_text("📝 Примечание? (или «-» если нет)")
    return P_NOTE


def _prod_summary(d):
    """Текст-подтверждение сохранённой записи производства."""
    total = (parse_num(d.get("фракция_0_1", 0)) + parse_num(d.get("фракция_1_2", 0))
             + parse_num(d.get("фракция_2_4", 0)) + parse_num(d.get("фракция_4_6", 0))
             + parse_num(d.get("фракция_6_8", 0)))
    bags = round(total / 30) if total else 0
    return (
        f"✅ *Производство сохранено!*\n\n"
        f"📅 Дата: {d.get('дата', '—')}\n"
        f"👤 Оператор: {d.get('фио', '—')}\n"
        f"⚖️ Вес шин: {d.get('вес_шин', 0)} кг\n"
        f"🛍 Мешки: {bags} шт (всего ÷ 30)\n\n"
        f"*Фракции (кг):*\n"
        f"  0-1: {d.get('фракция_0_1', 0)}\n"
        f"  1-2: {d.get('фракция_1_2', 0)}\n"
        f"  2-4: {d.get('фракция_2_4', 0)}\n"
        f"  4-6: {d.get('фракция_4_6', 0)}\n"
        f"  6-8: {d.get('фракция_6_8', 0)}\n"
        f"  Всего: {total} кг\n\n"
        f"🔩 Металлокорд: {d.get('металл_корд', 0)} кг"
    )


async def manual_prod_note(update: Update, context: ContextTypes.DEFAULT_TYPE):
    note = update.message.text.strip()
    context.user_data["prod"]["примечание"] = "" if note in ("-", "—") else note
    d = context.user_data["prod"]

    # NEW: при ручном вводе тоже предупреждаем о вероятном дубле (но сохраняем — это явное действие человека)
    dup_note = ""
    try:
        if find_duplicate_production(d):
            dup_note = "\n\n⚠️ Похоже, запись с этой датой и теми же фракциями уже есть — добавил ещё одну (ручной ввод)."
    except Exception as e:
        logger.error(f"dup check error (manual): {e}")

    save_production(d)
    await update.message.reply_text(_prod_summary(d) + dup_note, parse_mode="Markdown")
    context.user_data.pop("prod", None)
    return ConversationHandler.END


# ---------- Ручной ввод: РЕАЛИЗАЦИЯ ----------

async def manual_sale_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("⛔ Нет доступа. Обратитесь к администратору.")
        return ConversationHandler.END
    context.user_data["sale"] = {}
    await update.message.reply_text(
        "✍️ Ручной ввод реализации.\n\n"
        "Введи *дату* (дд.мм.гггг) или «сегодня».\n\n"
        "Для отмены — /cancel",
        parse_mode="Markdown"
    )
    return S_DATE


async def manual_sale_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    t = update.message.text.strip()
    if t.lower() == "сегодня":
        context.user_data["sale"]["дата"] = datetime.now().strftime("%d.%m.%Y")
    else:
        dt = parse_date_or_none(t)
        if dt is None:
            await update.message.reply_text(
                "⚠️ Не понял дату. Введи в виде ДД.ММ.ГГГГ (например 05.06.2026)\n"
                "или напиши «сегодня». Для отмены — /cancel"
            )
            return S_DATE
        context.user_data["sale"]["дата"] = date_to_str(dt)
    await update.message.reply_text("🏢 Покупатель?")
    return S_BUYER


async def manual_sale_buyer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["sale"]["покупатель"] = update.message.text.strip()
    kb = ReplyKeyboardMarkup([["Безналичный", "Наличный"]], resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text("💳 Тип расчёта? (Безналичный / Наличный)", reply_markup=kb)
    return S_PAYTYPE


async def manual_sale_paytype(update: Update, context: ContextTypes.DEFAULT_TYPE):
    val = update.message.text.strip().lower()
    if "нал" in val and "без" not in val:
        pay = "Наличный"
    elif "без" in val:
        pay = "Безналичный"
    else:
        await update.message.reply_text("⚠️ Выбери: Безналичный или Наличный")
        return S_PAYTYPE
    context.user_data["sale"]["тип_расчета"] = pay
    await update.message.reply_text(
        "📦 Фракция? (например: 2-4)\n_Доступны: 0-1, 1-2, 2-4, 4-6, 6-8_",
        parse_mode="Markdown"
    )
    return S_FRAC


async def manual_sale_frac(update: Update, context: ContextTypes.DEFAULT_TYPE):
    frac = update.message.text.strip()
    if not any(f in frac for f in FRACTIONS):
        await update.message.reply_text(
            "⚠️ Не распознал фракцию. Введи одну из: 0-1, 1-2, 2-4, 4-6, 6-8"
        )
        return S_FRAC
    context.user_data["sale"]["фракция"] = frac
    await update.message.reply_text("⚖️ Количество, кг?")
    return S_KG


async def manual_sale_kg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["sale"]["количество_кг"] = parse_num(update.message.text)
    await update.message.reply_text("💵 Цена за кг, тнг? (как в накладной)")
    return S_PRICE


async def manual_sale_price(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["sale"]["цена_за_кг"] = parse_num(update.message.text)
    await update.message.reply_text("💰 Сумма с НДС, тнг?")
    return S_SUM_VAT


async def manual_sale_sum_vat(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["sale"]["сумма_с_ндс"] = parse_num(update.message.text)
    await update.message.reply_text("💰 Сумма НДС, тнг? (или «-» если нет)")
    return S_VAT


async def manual_sale_vat(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["sale"]["сумма_ндс"] = parse_num(update.message.text)
    await update.message.reply_text("📝 Примечание? (или «-» если нет)")
    return S_NOTE


async def manual_sale_note(update: Update, context: ContextTypes.DEFAULT_TYPE):
    note = update.message.text.strip()
    context.user_data["sale"]["примечание"] = "" if note in ("-", "—") else note
    d = context.user_data["sale"]
    save_sale(d)

    report_line = ""
    try:
        ok, msg = write_sale_to_report(d)
        report_line = ("\n\n🧾 Отчёт: " + msg) if ok else ("\n\n⚠️ Отчёт: " + msg)
    except Exception as e:
        logger.error(f"Report write error: {e}")
        report_line = "\n\n⚠️ Отчёт: не удалось записать (склад сохранён)."

    qty_kg = parse_num(d.get("количество_кг", 0))
    reply = (
        f"✅ *Реализация сохранена!*\n\n"
        f"📅 Дата: {d.get('дата', '—')}\n"
        f"🏢 Покупатель: {d.get('покупатель', '—')}\n"
        f"💳 Тип расчёта: {d.get('тип_расчета', '—')}\n"
        f"📦 Фракция: {d.get('фракция', '—')}\n"
        f"⚖️ Количество: {qty_kg} кг\n"
        f"💵 Цена за кг: {d.get('цена_за_кг', 0)} тнг\n"
        f"💰 Сумма с НДС: {d.get('сумма_с_ндс', 0)} тнг\n"
        f"💰 Сумма НДС: {d.get('сумма_ндс', 0)} тнг" + report_line
    )
    await update.message.reply_text(reply, parse_mode="Markdown")
    context.user_data.pop("sale", None)
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("prod", None)
    context.user_data.pop("sale", None)
    context.user_data.pop("pending_prod", None)
    context.user_data.pop("pending_sale", None)
    context.user_data.pop("pending_photo", None)
    context.user_data.pop("exp", None)
    await update.message.reply_text("❌ Отменено.")
    return ConversationHandler.END


# ---------- Фото и прочий текст ----------

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("⛔ Нет доступа. Обратитесь к администратору.")
        return
    text = update.message.text
    if text == "📸 Производство":
        context.user_data["photo_type"] = "production"
        await update.message.reply_text("📸 Отправь фото отчёта по производству крошки.")
    elif text == "📄 Реализация":
        context.user_data["photo_type"] = "sale"
        await update.message.reply_text("📄 Отправь фото накладной на реализацию.")
    else:
        await update.message.reply_text("Используй кнопки меню или команды /ostatok, /last")


async def _save_sale_from_photo(update, data):
    """Сохранение реализации с фото + попытка записи в отчёт (вынесено для чистоты)."""
    data.setdefault("тип_расчета", "")
    save_sale(data)

    report_line = ""
    try:
        ok, msg = write_sale_to_report(data)
        report_line = ("\n🧾 Отчёт: " + msg) if ok else ("\n⚠️ Отчёт: " + msg)
    except Exception as e:
        logger.error(f"Report write error: {e}")
        report_line = "\n⚠️ Отчёт: не удалось записать (склад сохранён)."

    qty_kg = parse_num(data.get("количество_кг", 0))
    if not qty_kg and data.get("количество_т"):
        qty_kg = parse_num(data.get("количество_т", 0)) * 1000
    reply = (
        f"✅ *Реализация сохранена!*\n\n"
        f"📅 Дата: {data.get('дата', '—')}\n"
        f"🏢 Покупатель: {data.get('покупатель', '—')}\n"
        f"💳 Тип расчёта: {data.get('тип_расчета') or '— (уточни вручную)'}\n"
        f"📦 Фракция: {data.get('фракция', '—')}\n"
        f"⚖️ Количество: {qty_kg:.0f} кг\n"
        f"💵 Цена за кг: {data.get('цена_за_кг', 0)} тнг\n"
        f"💰 Сумма с НДС: {data.get('сумма_с_ндс', 0)} тнг\n"
        f"💰 Сумма НДС: {data.get('сумма_ндс', 0)} тнг" + report_line
    )
    await update.message.reply_text(reply, parse_mode="Markdown")


async def photo_sale_paytype(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Доспрос нал/безнал после фото-реализации, затем запись в склад и отчёт."""
    val = update.message.text.strip().lower()
    if "нал" in val and "без" not in val:
        pay = "Наличный"
    elif "без" in val:
        pay = "Безналичный"
    else:
        await update.message.reply_text("⚠️ Выбери: Безналичный или Наличный")
        return PHOTO_SALE_PAY
    data = context.user_data.get("pending_sale", {})
    data["тип_расчета"] = pay
    await _save_sale_from_photo(update, data)
    context.user_data.pop("pending_sale", None)
    context.user_data["photo_type"] = "production"
    return ConversationHandler.END


async def detect_doc_type(image_bytes: bytes):
    """Определяет по фото: накладная-реализация или отчёт производства."""
    image_b64 = base64.b64encode(image_bytes).decode()
    prompt = (
        "Определи тип документа на фото. Это одно из двух:\n"
        "- НАКЛАДНАЯ на отпуск/реализацию (есть покупатель, цены, суммы, НДС) — это продажа;\n"
        "- ОТЧЁТ ПРОИЗВОДСТВА крошки (смена, оператор, вес шин, фракции в кг, мешки).\n"
        'Верни ТОЛЬКО JSON: {"тип":"реализация"} или {"тип":"производство"}.'
    )
    try:
        res = await call_claude(image_b64, prompt)
        t = str(res.get("тип", "")).strip().lower() if isinstance(res, dict) else ""
    except Exception as e:
        logger.error(f"detect_doc_type error: {e}")
        t = ""
    return "sale" if "реализ" in t else "production"


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Приём фото: определяем тип (реализация/производство) и просим подтвердить."""
    if not is_allowed(update):
        await update.message.reply_text("⛔ Нет доступа. Обратитесь к администратору.")
        return ConversationHandler.END
    await update.message.reply_text("📸 Получил фото, определяю тип документа...")
    try:
        photo = update.message.photo[-1]
        file = await context.bot.get_file(photo.file_id)
        image_bytes = bytes(await file.download_as_bytearray())
        doc_type = await detect_doc_type(image_bytes)
        context.user_data["pending_photo"] = {"file_id": photo.file_id, "type": doc_type}
        if doc_type == "sale":
            label, other = "📄 РЕАЛИЗАЦИЯ (накладная)", "🔁 Нет, это производство"
        else:
            label, other = "📸 ПРОИЗВОДСТВО (отчёт)", "🔁 Нет, это реализация"
        kb = ReplyKeyboardMarkup([["✅ Да, верно", other]],
                                 resize_keyboard=True, one_time_keyboard=True)
        await update.message.reply_text(
            f"Определил: {label}\n\nВерно? Подтверди — распознаю и сохраню.",
            reply_markup=kb,
        )
        return PHOTO_TYPE_CONFIRM
    except Exception as e:
        logger.error(f"handle_photo detect error: {e}")
        await update.message.reply_text("❌ Не смог обработать фото. Попробуй ещё раз.")
        return ConversationHandler.END


async def photo_type_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Подтверждение типа: Да — сохраняем; Нет — берём другой тип."""
    ans = update.message.text.strip().lower()
    pending = context.user_data.get("pending_photo")
    if not pending:
        await update.message.reply_text("Пришли фото заново, пожалуйста.")
        return ConversationHandler.END
    if "да" in ans or "верно" in ans:
        photo_type = pending["type"]
    elif "нет" in ans or "производ" in ans or "реализ" in ans:
        photo_type = "production" if pending["type"] == "sale" else "sale"
    else:
        await update.message.reply_text("Нажми «✅ Да, верно» или «🔁 Нет…».")
        return PHOTO_TYPE_CONFIRM
    context.user_data.pop("pending_photo", None)
    await update.message.reply_text("🔎 Принято, распознаю данные...")
    try:
        file = await context.bot.get_file(pending["file_id"])
        image_bytes = bytes(await file.download_as_bytearray())
    except Exception as e:
        logger.error(f"photo_type_confirm download error: {e}")
        await update.message.reply_text("❌ Не смог скачать фото. Пришли заново.")
        return ConversationHandler.END
    return await _dispatch_photo(update, context, photo_type, image_bytes)


async def _dispatch_photo(update, context, photo_type, image_bytes):
    """Распознаёт и сохраняет фото по подтверждённому типу."""
    try:
        if photo_type == "sale":
            data = await recognize_sale(image_bytes)
            logger.info(f"Sale data: {data}")
            # тип расчёта (нал/безнал) на накладной обычно не виден — нормализуем распознанное
            pay_raw = str(data.get("тип_расчета", "")).strip().lower()
            if "нал" in pay_raw and "без" not in pay_raw:
                data["тип_расчета"] = "Наличный"
            elif "без" in pay_raw:
                data["тип_расчета"] = "Безналичный"
            else:
                data["тип_расчета"] = ""
            if not data["тип_расчета"]:
                # доспрашиваем, чтобы продажа попала и в отчёт «Реализация 2026г»
                context.user_data["pending_sale"] = data
                kb = ReplyKeyboardMarkup([["Безналичный", "Наличный"]],
                                         resize_keyboard=True, one_time_keyboard=True)
                await update.message.reply_text(
                    "📄 Накладную распознал. Уточни тип расчёта, чтобы продажа попала и в отчёт:",
                    reply_markup=kb,
                )
                return PHOTO_SALE_PAY
            await _save_sale_from_photo(update, data)
            context.user_data["photo_type"] = "production"
            return ConversationHandler.END

        # --- производство ---
        data = await recognize_production(image_bytes)
        logger.info(f"Production data: {data}")
        records = data if isinstance(data, list) else [data]

        # Проверяем дату КАЖДОЙ записи. Claude должен был отдать дд.мм.гггг.
        # Если дата не читается — НЕ пишем молча, а просим ввести вручную.
        bad_date = []
        for r in records:
            dt = parse_date_or_none(r.get("дата", ""))
            if dt is None:
                bad_date.append(r)
            else:
                r["дата"] = date_to_str(dt)  # приводим к единому виду дд.мм.гггг

        if bad_date:
            await update.message.reply_text(
                "⚠️ Не смог уверенно прочитать ДАТУ на фото.\n\n"
                "Чтобы не записать неверно, внеси эту смену через «✍️ Ввод производства» "
                "(там дату введёшь вручную), или переснимай фото так, чтобы дата была чёткой "
                "и с точками (например 05.06.2026).\n\n"
                "В склад с этого фото ничего не записал."
            )
            return ConversationHandler.END

        # Сверка: сумма по фракциям должна совпасть с «всего», написанным на бланке
        mismatch = []
        for r in records:
            written = parse_num(r.get("всего_итог", 0))
            calc = sum(parse_num(r.get(fk, 0)) for fk in
                       ("фракция_0_1", "фракция_1_2", "фракция_2_4", "фракция_4_6", "фракция_6_8"))
            if written and abs(calc - written) > 1:
                mismatch.append((calc, written))
        if mismatch:
            calc, written = mismatch[0]
            await update.message.reply_text(
                "⚠️ Расхождение по сумме крошки.\n\n"
                f"На бланке «всего»: {written:.0f} кг\n"
                f"По распознанным фракциям: {calc:.0f} кг\n\n"
                "Возможно, цифра распозналась неверно. Проверь фото и внеси смену вручную "
                "через «✍️ Ввод производства», или переснимай чётче.\n\n"
                "В склад с этого фото ничего не записал."
            )
            return ConversationHandler.END

        # Делим на «чистые» (сразу пишем) и «вероятные дубли» (спросим)
        clean, dups = [], []
        for r in records:
            try:
                if find_duplicate_production(r):
                    dups.append(r)
                else:
                    clean.append(r)
            except Exception as e:
                logger.error(f"dup check error (photo): {e}")
                clean.append(r)  # если проверка упала — не теряем данные, пишем

        # Чистые записи сохраняем сразу
        for item in clean:
            save_production(item)

        # Если есть подозрение на дубль — спрашиваем по первому, остальные дубли держим в очереди
        if dups:
            context.user_data["pending_prod"] = dups
            first_dup = dups[0]
            sig_total = sum(_fractions_signature(first_dup))
            saved_note = ""
            if clean:
                saved_note = f"\n\n(Заодно сохранил новых записей: {len(clean)}.)"
            kb = ReplyKeyboardMarkup([["Да, добавить", "Нет, пропустить"]],
                                     resize_keyboard=True, one_time_keyboard=True)
            await update.message.reply_text(
                f"⚠️ Похоже, эта запись УЖЕ ЕСТЬ в складе:\n\n"
                f"📅 Дата: {first_dup.get('дата', '—')}\n"
                f"👤 Оператор: {first_dup.get('фио', '—')}\n"
                f"⚖️ Всего крошки: {sig_total:.0f} кг\n\n"
                f"Добавить её всё равно?" + saved_note,
                reply_markup=kb,
                parse_mode="Markdown"
            )
            return PHOTO_PROD_CONFIRM

        # Дублей нет — обычный отчёт
        if not clean:
            await update.message.reply_text("⚠️ С фото не удалось получить данные. Попробуй ещё раз.")
            return ConversationHandler.END

        first = clean[-1]
        await update.message.reply_text(
            _prod_summary(first).replace("сохранено!", f"сохранено! (новых: {len(clean)})"),
            parse_mode="Markdown"
        )
        return ConversationHandler.END

    except Exception as e:
        logger.error(f"Error: {e}")
        await update.message.reply_text(
            "❌ Не смог распознать фото. Попробуй:\n"
            "• Сделать фото чётче\n"
            "• Хорошее освещение\n"
            "• Всё в кадре"
        )
        return ConversationHandler.END


async def photo_prod_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """NEW: ответ Да/Нет на вопрос о дубле фото-производства."""
    ans = update.message.text.strip().lower()
    dups = context.user_data.get("pending_prod", [])
    context.user_data.pop("pending_prod", None)

    if not dups:
        await update.message.reply_text("Нечего подтверждать. Отправь фото заново.")
        return ConversationHandler.END

    if "да" in ans:
        for item in dups:
            save_production(item)
        await update.message.reply_text(
            f"✅ Добавил, несмотря на совпадение. Записей: {len(dups)}.\n"
            f"Проверь остаток: /ostatok"
        )
    else:
        await update.message.reply_text(
            "👍 Пропустил — дубль в склад не попал."
        )
    return ConversationHandler.END


# ---------- Автобэкап данных в закрытый Telegram-канал ----------

def _build_backup_xlsx():
    """Собирает один .xlsx: листы склада (Производство, Реализация) + вкладки отчёта реализации.
    Возвращает (bytes_файла, counts) — counts это число строк данных по ключевым листам."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    counts = {}
    book = get_spreadsheet()
    for name in ("Производство", "Реализация"):
        try:
            rows = book.worksheet(name).get_all_values()
        except Exception as e:
            logger.error(f"backup read {name}: {e}")
            rows = [["(не удалось прочитать лист)"]]
        sheet = wb.create_sheet(name[:31])
        for row in rows:
            sheet.append(row)
        counts[name] = max(0, len(rows) - 1)
    try:
        rep = get_report_spreadsheet()
        for ws in rep.worksheets():
            rows = ws.get_all_values()
            sheet = wb.create_sheet(("Отчёт-" + ws.title)[:31])
            for row in rows:
                sheet.append(row)
    except Exception as e:
        logger.error(f"backup read report: {e}")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), counts


async def do_backup(bot):
    """Собирает и отправляет бэкап в закрытый канал. Возвращает 'ok' или текст ошибки."""
    if not BACKUP_CHANNEL_ID:
        logger.warning("BACKUP_CHANNEL_ID не задан — бэкап пропущен")
        return "не задан BACKUP_CHANNEL_ID"
    data, counts = await asyncio.to_thread(_build_backup_xlsx)
    global _last_backup_counts
    warn = ""
    for key in ("Производство", "Реализация"):
        old = _last_backup_counts.get(key)
        new = counts.get(key, 0)
        if old is not None and (old - new) > 2:
            warn += f"\n⚠️ В «{key}» строк стало МЕНЬШЕ: было {old}, стало {new}. Проверь, не удалили ли случайно!"
    _last_backup_counts = counts
    now = datetime.now(KZ_TZ)
    fname = "edil_backup_" + now.strftime("%Y-%m-%d_%H%M") + ".xlsx"
    caption = ("🗄 Бэкап Едил · " + now.strftime("%d.%m.%Y %H:%M")
               + f"\nПроизводство: {counts.get('Производство', 0)} · Реализация: {counts.get('Реализация', 0)} строк")
    if warn:
        caption += "\n" + warn
    await bot.send_document(
        chat_id=int(BACKUP_CHANNEL_ID),
        document=InputFile(io.BytesIO(data), filename=fname),
        caption=caption,
    )
    return "ok"


async def _backup_loop(application):
    """Раз в сутки в BACKUP_HOUR по местному времени (UTC+5) делает бэкап."""
    await asyncio.sleep(10)
    while True:
        now = datetime.now(KZ_TZ)
        target = now.replace(hour=BACKUP_HOUR, minute=0, second=0, microsecond=0)
        if target <= now:
            target += timedelta(days=1)
        await asyncio.sleep(max(1, (target - now).total_seconds()))
        try:
            await do_backup(application.bot)
        except Exception as e:
            logger.error(f"scheduled backup error: {e}")


async def _post_init(application):
    application.create_task(_backup_loop(application))


async def cmd_backup(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Ручной запуск бэкапа для проверки. Только для разрешённых пользователей."""
    if not is_allowed(update):
        await update.message.reply_text("⛔ Нет доступа.")
        return
    await update.message.reply_text("🗄 Делаю бэкап, пара секунд...")
    try:
        res = await do_backup(context.bot)
    except Exception as e:
        logger.error(f"manual backup error: {e}")
        await update.message.reply_text(f"❌ Ошибка бэкапа: {e}")
        return
    if res == "ok":
        await update.message.reply_text("✅ Бэкап отправлен в канал.")
    else:
        await update.message.reply_text(f"⚠️ {res}")


# ---------- Журнал «Расходы»: ввод расходов компании ----------

SHEET_EXPENSES = "Расходы"
EXPENSE_HEADERS = ["Дата", "Сумма", "Группа", "Категория", "Нал/Безнал", "Источник", "Контрагент", "Примечание"]
EXPENSE_GROUPS = {
    "Производство": [
        "Материалы и запчасти", "Упаковка (мешки)", "Электроэнергия", "ГСМ",
        "ФОТ производственный", "Транспорт и логистика", "Лизинг (вознаграждение)",
        "Амортизация производства",
    ],
    "Администрация": [
        "ФОТ административный", "Налоги", "Услуги банка", "Страхование", "Аудит",
        "Прочие административные", "Амортизация административная",
    ],
    "Капзатраты": [
        "Основные средства", "Лизинг (основной долг)", "Прочие капзатраты",
    ],
}
EXPENSE_SOURCES = ["Евразийский", "БЦК", "Касса 1", "Касса 2", "Неденежный"]


def get_expenses_ws():
    return get_worksheet(SHEET_EXPENSES, EXPENSE_HEADERS)


def _source_to_paytype(src):
    if src in ("Евразийский", "БЦК"):
        return "Безналичный"
    if src in ("Касса 1", "Касса 2"):
        return "Наличный"
    return "Неденежный"


def save_expense(data):
    ws = get_expenses_ws()
    src = data.get("источник", "")
    row = [
        data.get("дата", datetime.now().strftime("%d.%m.%Y")),
        parse_num(data.get("сумма", 0)),
        data.get("группа", ""),
        data.get("категория", ""),
        _source_to_paytype(src),
        src,
        data.get("контрагент", ""),
        data.get("примечание", ""),
    ]
    ws.append_row(row, value_input_option="USER_ENTERED")


def _exp_summary(d, saved=True):
    head = "✅ *Расход сохранён!*" if saved else "Проверь расход перед сохранением:"
    return (
        f"{head}\n\n"
        f"📅 Дата: {d.get('дата', '—')}\n"
        f"💵 Сумма: {d.get('сумма', 0)} тнг\n"
        f"📂 Группа: {d.get('группа', '—')}\n"
        f"🏷 Категория: {d.get('категория', '—')}\n"
        f"🏦 Источник: {d.get('источник', '—')}\n"
        f"🏢 Контрагент: {d.get('контрагент') or '—'}\n"
        f"📝 Примечание: {d.get('примечание') or '—'}"
    )


async def manual_exp_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("⛔ Нет доступа.")
        return ConversationHandler.END
    context.user_data["exp"] = {}
    await update.message.reply_text("💸 Внесём расход. Пришли ФОТО счёта/чека — распознаю сам, или введи дату вручную (дд.мм.гггг или «сегодня»).")
    return E_DATE


async def manual_exp_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    t = update.message.text.strip().lower()
    if t in ("сегодня", "today"):
        context.user_data["exp"]["дата"] = datetime.now().strftime("%d.%m.%Y")
    else:
        dt = parse_date_or_none(update.message.text)
        if dt is None:
            await update.message.reply_text("⚠️ Не понял дату. Введи в виде 08.06.2026 или «сегодня».")
            return E_DATE
        context.user_data["exp"]["дата"] = date_to_str(dt)
    await update.message.reply_text("💵 Сумма расхода, тенге?")
    return E_AMOUNT


async def manual_exp_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    amt = parse_num(update.message.text)
    if not amt:
        await update.message.reply_text("⚠️ Введи сумму числом, например 150000.")
        return E_AMOUNT
    context.user_data["exp"]["сумма"] = amt
    kb = ReplyKeyboardMarkup([["Производство"], ["Администрация"], ["Капзатраты"]],
                             resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text("📂 Группа расхода?", reply_markup=kb)
    return E_GROUP


async def manual_exp_group(update: Update, context: ContextTypes.DEFAULT_TYPE):
    grp = update.message.text.strip()
    if grp not in EXPENSE_GROUPS:
        await update.message.reply_text("⚠️ Выбери кнопкой: Производство, Администрация или Капзатраты.")
        return E_GROUP
    context.user_data["exp"]["группа"] = grp
    kb = ReplyKeyboardMarkup([[c] for c in EXPENSE_GROUPS[grp]],
                             resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text("🏷 Категория?", reply_markup=kb)
    return E_CATEGORY


async def manual_exp_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    cat = update.message.text.strip()
    grp = context.user_data["exp"].get("группа", "")
    if cat not in EXPENSE_GROUPS.get(grp, []):
        await update.message.reply_text("⚠️ Выбери категорию кнопкой из списка.")
        return E_CATEGORY
    context.user_data["exp"]["категория"] = cat
    if context.user_data["exp"].get("источник"):
        kb = ReplyKeyboardMarkup([["✅ Да, сохранить", "❌ Отмена"]],
                                 resize_keyboard=True, one_time_keyboard=True)
        await update.message.reply_text(_exp_summary(context.user_data["exp"], saved=False),
                                        reply_markup=kb, parse_mode="Markdown")
        return E_CONFIRM
    kb = ReplyKeyboardMarkup([["Евразийский", "БЦК"], ["Касса 1", "Касса 2"], ["Неденежный"]],
                             resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text("🏦 Источник (откуда оплата)?", reply_markup=kb)
    return E_SOURCE


async def manual_exp_source(update: Update, context: ContextTypes.DEFAULT_TYPE):
    src = update.message.text.strip()
    if src not in EXPENSE_SOURCES:
        await update.message.reply_text("⚠️ Выбери источник кнопкой.")
        return E_SOURCE
    context.user_data["exp"]["источник"] = src
    if context.user_data["exp"].get("_photo"):
        kb = ReplyKeyboardMarkup([["✅ Да, сохранить", "❌ Отмена"]],
                                 resize_keyboard=True, one_time_keyboard=True)
        await update.message.reply_text(_exp_summary(context.user_data["exp"], saved=False),
                                        reply_markup=kb, parse_mode="Markdown")
        return E_CONFIRM
    await update.message.reply_text("🏢 Контрагент (кому платим)? Или «-», если не нужно.")
    return E_CONTRAGENT


async def manual_exp_contragent(update: Update, context: ContextTypes.DEFAULT_TYPE):
    t = update.message.text.strip()
    context.user_data["exp"]["контрагент"] = "" if t in ("-", "—") else t
    await update.message.reply_text("📝 Примечание? Или «-», если нет.")
    return E_NOTE


async def manual_exp_note(update: Update, context: ContextTypes.DEFAULT_TYPE):
    t = update.message.text.strip()
    context.user_data["exp"]["примечание"] = "" if t in ("-", "—") else t
    kb = ReplyKeyboardMarkup([["✅ Да, сохранить", "❌ Отмена"]],
                             resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text(_exp_summary(context.user_data["exp"], saved=False),
                                    reply_markup=kb, parse_mode="Markdown")
    return E_CONFIRM


async def manual_exp_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ans = update.message.text.strip().lower()
    if "да" in ans or "сохран" in ans:
        d = context.user_data["exp"]
        save_expense(d)
        await update.message.reply_text(_exp_summary(d, saved=True), parse_mode="Markdown")
        context.user_data.pop("exp", None)
        return ConversationHandler.END
    if "отмена" in ans or "нет" in ans:
        context.user_data.pop("exp", None)
        await update.message.reply_text("❌ Отменено, расход не сохранён.")
        return ConversationHandler.END
    await update.message.reply_text("Нажми «✅ Да, сохранить» или «❌ Отмена».")
    return E_CONFIRM


async def recognize_expense(image_bytes: bytes):
    """Распознаёт расходный документ (счёт/чек/накладную на покупку)."""
    image_b64 = base64.b64encode(image_bytes).decode()
    prompt = (
        "Это фото счёта, чека, платёжного поручения или накладной на РАСХОД (компания оплачивает товар или услугу).\n"
        "Верни ТОЛЬКО JSON без markdown:\n"
        '{"дата":"дд.мм.гггг","контрагент":"кому платим (поставщик/исполнитель)","сумма":0,"банк":"банк плательщика, если виден"}\n'
        "Сумма — итог к оплате. Контрагент — ПОЛУЧАТЕЛЬ платежа. Банк — банк ПЛАТЕЛЬЩИКА (наш счёт, откуда деньги), если виден. Если что-то не читается — 0 или пусто."
    )
    return await call_claude(image_b64, prompt)


async def manual_exp_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Фото счёта/чека внутри «Ввод расхода»: распознаём сумму/контрагента/дату."""
    await update.message.reply_text("📸 Получил фото, распознаю расход...")
    try:
        photo = update.message.photo[-1]
        file = await context.bot.get_file(photo.file_id)
        image_bytes = bytes(await file.download_as_bytearray())
        data = await recognize_expense(image_bytes)
        logger.info(f"Expense photo data: {data}")
    except Exception as e:
        logger.error(f"manual_exp_photo error: {e}")
        await update.message.reply_text("❌ Не смог распознать фото. Введи дату вручную (дд.мм.гггг) или пришли фото чётче.")
        return E_DATE
    dt = parse_date_or_none(data.get("дата", ""))
    context.user_data["exp"] = {
        "_photo": True,
        "дата": date_to_str(dt) if dt else datetime.now().strftime("%d.%m.%Y"),
        "сумма": parse_num(data.get("сумма", 0)),
        "контрагент": str(data.get("контрагент", "")).strip(),
    }
    bank_raw = str(data.get("банк", "")).lower()
    if "евраз" in bank_raw:
        context.user_data["exp"]["источник"] = "Евразийский"
    elif "центркредит" in bank_raw or "бцк" in bank_raw or "centercredit" in bank_raw:
        context.user_data["exp"]["источник"] = "БЦК"
    d = context.user_data["exp"]
    kb = ReplyKeyboardMarkup([["Производство"], ["Администрация"], ["Капзатраты"]],
                             resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text(
        f"💸 Распознал:\n💵 Сумма: {d['сумма']} тнг\n🏢 Контрагент: {d['контрагент'] or '—'}\n📅 Дата: {d['дата']}\n\nВыбери группу расхода:",
        reply_markup=kb,
    )
    return E_GROUP


def main():
    app = Application.builder().token(BOT_TOKEN).post_init(_post_init).build()

    prod_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^✍️ Ввод производства$"), manual_prod_start)],
        states={
            P_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_date)],
            P_FIO: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_fio)],
            P_TIRES: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_tires)],
            P_THREAD: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_thread)],
            P_F01: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_f01)],
            P_F12: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_f12)],
            P_F24: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_f24)],
            P_F46: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_f46)],
            P_F68: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_f68)],
            P_CORD: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_cord)],
            P_NOTE: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_note)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    sale_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^✍️ Ввод реализации$"), manual_sale_start)],
        states={
            S_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_date)],
            S_BUYER: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_buyer)],
            S_PAYTYPE: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_paytype)],
            S_FRAC: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_frac)],
            S_KG: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_kg)],
            S_PRICE: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_price)],
            S_SUM_VAT: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_sum_vat)],
            S_VAT: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_vat)],
            S_NOTE: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_note)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    # NEW: диалог фото-производства с подтверждением дубля.
    # entry_point — приход фото; если дубль, переходим в PHOTO_PROD_CONFIRM и ждём Да/Нет.
    photo_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.PHOTO, handle_photo)],
        states={
            PHOTO_TYPE_CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, photo_type_confirm)],
            PHOTO_PROD_CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, photo_prod_confirm)],
            PHOTO_SALE_PAY: [MessageHandler(filters.TEXT & ~filters.COMMAND, photo_sale_paytype)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("ostatok", ostatok))
    app.add_handler(CommandHandler("last", last_records))
    app.add_handler(CommandHandler("cancel", cancel))
    app.add_handler(CommandHandler("backup", cmd_backup))
    exp_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^💸 Ввод расхода$"), manual_exp_start)],
        states={
            E_DATE: [MessageHandler(filters.PHOTO, manual_exp_photo), MessageHandler(filters.TEXT & ~filters.COMMAND, manual_exp_date)],
            E_AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_exp_amount)],
            E_GROUP: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_exp_group)],
            E_CATEGORY: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_exp_category)],
            E_SOURCE: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_exp_source)],
            E_CONTRAGENT: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_exp_contragent)],
            E_NOTE: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_exp_note)],
            E_CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_exp_confirm)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(prod_conv)
    app.add_handler(sale_conv)
    app.add_handler(exp_conv)
    app.add_handler(photo_conv)  # после exp_conv: фото в режиме «Ввод расхода» ловит exp_conv
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    if WEBHOOK_URL:
        logger.info(f"Starting webhook on port {PORT}")
        app.run_webhook(
            listen="0.0.0.0",
            port=PORT,
            webhook_url=f"{WEBHOOK_URL}/{BOT_TOKEN}",
            url_path=BOT_TOKEN,
            drop_pending_updates=True
        )
    else:
        logger.info("Starting polling...")
        app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
