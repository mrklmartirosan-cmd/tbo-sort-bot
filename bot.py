import os
import json
import base64
import logging
import re
from datetime import datetime
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, filters,
    ContextTypes, ConversationHandler
)
import httpx
import gspread

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
WEBHOOK_URL = os.environ.get("WEBHOOK_URL", "")
PORT = int(os.environ.get("PORT", 8080))

# --- Google Sheets ---
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
REPORT_SPREADSHEET_ID = os.environ.get("REPORT_SPREADSHEET_ID", "")  # NEW: файл "Реализация 2026г"
GOOGLE_CREDENTIALS = os.environ.get("GOOGLE_CREDENTIALS", "")  # содержимое JSON-ключа целиком

SHEET_PRODUCTION = "Производство"
SHEET_SALES = "Реализация"

# NEW: список разрешённых Telegram ID (закрытый доступ).
# Пустой список = пускать всех (страховка, чтобы не заблокировать себя).
_allowed_raw = os.environ.get("ALLOWED_USERS", "")
ALLOWED_USERS = {int(x) for x in _allowed_raw.replace(" ", "").split(",") if x.strip().isdigit()}

# NEW: настройки отчёта реализации ---------------------------------
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
# CHANGED: добавлено состояние S_PRICE (цена за кг)
(S_DATE, S_BUYER, S_PAYTYPE, S_FRAC, S_KG, S_PRICE, S_SUM_VAT, S_VAT, S_NOTE) = range(12, 21)

FRACTIONS = ["0-1", "1-2", "2-4", "4-6", "6-8"]

# Заголовки листов склада (порядок колонок = порядок записи)
PROD_HEADERS = ["Дата", "ФИО оператора", "Вес шин кг", "Мешки шт", "Нитки",
                "Фракция 0-1", "Фракция 1-2", "Фракция 2-4", "Фракция 4-6", "Фракция 6-8",
                "Всего крошки кг", "Металлокорд кг", "Примечание"]
# CHANGED: добавлена "Цена за кг" между "Количество кг" и "Сумма с НДС"
SALES_HEADERS = ["Дата", "Покупатель", "Тип расчёта", "Фракция", "Количество кг",
                 "Цена за кг", "Сумма с НДС", "Сумма НДС", "Примечание"]

_gc = None
_spreadsheet = None
_report_spreadsheet = None  # NEW


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


# ---------- NEW: контроль доступа ----------

def is_allowed(update: Update) -> bool:
    """True, если пользователь в списке разрешённых.
    Если ALLOWED_USERS пуст — пускаем всех (страховка от самоблокировки)."""
    if not ALLOWED_USERS:
        return True
    user = update.effective_user
    return bool(user and user.id in ALLOWED_USERS)


# ---------- Google Sheets: подключение и доступ к листам ----------

def _get_client():
    """Авторизация служебного аккаунта (один раз)."""
    global _gc
    if _gc is None:
        if not GOOGLE_CREDENTIALS:
            raise RuntimeError("Не задан GOOGLE_CREDENTIALS в переменных окружения.")
        creds_dict = json.loads(GOOGLE_CREDENTIALS)
        _gc = gspread.service_account_from_dict(creds_dict)
    return _gc


def get_spreadsheet():
    """Ленивое подключение к таблице-складу. Кэшируем, чтобы не авторизоваться каждый раз."""
    global _spreadsheet
    if _spreadsheet is not None:
        return _spreadsheet
    if not SPREADSHEET_ID:
        raise RuntimeError("Не задан SPREADSHEET_ID в переменных окружения.")
    _spreadsheet = _get_client().open_by_key(SPREADSHEET_ID)
    return _spreadsheet


def get_report_spreadsheet():
    """NEW: ленивое подключение к файлу отчёта 'Реализация 2026г'."""
    global _report_spreadsheet
    if _report_spreadsheet is not None:
        return _report_spreadsheet
    if not REPORT_SPREADSHEET_ID:
        raise RuntimeError("Не задан REPORT_SPREADSHEET_ID в переменных окружения.")
    _report_spreadsheet = _get_client().open_by_key(REPORT_SPREADSHEET_ID)
    return _report_spreadsheet


def get_worksheet(title, headers):
    """Возвращает лист по имени. Если листа нет — создаёт и пишет шапку.
    Если лист пустой — пишет шапку."""
    sh = get_spreadsheet()
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=1000, cols=max(len(headers), 12))
        ws.append_row(headers, value_input_option="USER_ENTERED")
        return ws
    # Лист есть — проверим, есть ли шапка
    first = ws.row_values(1)
    if not first:
        ws.append_row(headers, value_input_option="USER_ENTERED")
    return ws


def get_production_ws():
    return get_worksheet(SHEET_PRODUCTION, PROD_HEADERS)


def get_sales_ws():
    return get_worksheet(SHEET_SALES, SALES_HEADERS)


# ---------- Чтение/запись данных ----------

def get_existing_dates_production():
    ws = get_production_ws()
    dates = set()
    col = ws.col_values(1)  # колонка "Дата"
    for val in col[1:]:  # пропускаем шапку
        if val:
            dates.add(str(val).strip())
    return dates


def save_production(data):
    ws = get_production_ws()
    f01 = parse_num(data.get("фракция_0_1", 0))
    f12 = parse_num(data.get("фракция_1_2", 0))
    f24 = parse_num(data.get("фракция_2_4", 0))
    f46 = parse_num(data.get("фракция_4_6", 0))
    f68 = parse_num(data.get("фракция_6_8", 0))
    total = f01 + f12 + f24 + f46 + f68
    row = [
        data.get("дата", datetime.now().strftime("%d.%m.%Y")),
        data.get("фио", ""),
        parse_num(data.get("вес_шин", 0)),
        parse_num(data.get("мешки", 0)),
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
        parse_num(data.get("цена_за_кг", 0)),  # CHANGED: новая колонка "Цена за кг"
        parse_num(data.get("сумма_с_ндс", 0)),
        parse_num(data.get("сумма_ндс", 0)),
        data.get("примечание", ""),
    ]
    ws.append_row(row, value_input_option="USER_ENTERED")


# ---------- NEW: запись реализации в отчёт "Реализация 2026г" ----------

def _normalize_fraction(frac_text):
    """Из произвольного текста фракции выделяет канон '0-1'/'1-2'/.../'6-8'."""
    for f in FRACTIONS:
        if f in str(frac_text):
            return f
    return None


def _parse_sale_date(date_text):
    """Пытается распарсить дату 'дд.мм.гггг'. Возвращает datetime или None."""
    s = str(date_text).strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return None


def _get_or_create_month_sheet(sh, dt):
    """Находит вкладку месяца ('июнь 2026'); если нет — копирует ШАБЛОН и переименовывает."""
    title = f"{RU_MONTHS[dt.month]} {dt.year}"
    try:
        return sh.worksheet(title)
    except gspread.WorksheetNotFound:
        pass
    # Листа нет — создаём из шаблона
    template = sh.worksheet(REPORT_TEMPLATE)  # бросит WorksheetNotFound, если шаблона нет
    new_ws = template.duplicate(new_sheet_name=title)
    return new_ws


def write_sale_to_report(data):
    """Дописывает строку реализации в отчёт. Возвращает (ok: bool, msg: str).
    Не бросает наружу — ошибки ловит вызывающий код по флагу."""
    pay = str(data.get("тип_расчета", "")).strip()
    if pay not in REPORT_BLOCKS:
        return False, f"тип расчёта не распознан ('{pay}') — строка в отчёт не добавлена"

    frac = _normalize_fraction(data.get("фракция", ""))
    if frac is None:
        return False, "фракция не распознана — строка в отчёт не добавлена"

    dt = _parse_sale_date(data.get("дата", ""))
    if dt is None:
        return False, "дата не распознана — строка в отчёт не добавлена"

    sh = get_report_spreadsheet()
    ws = _get_or_create_month_sheet(sh, dt)

    start_row, end_row = REPORT_BLOCKS[pay]
    # Ищем первую свободную строку блока (пустой столбец B = Дата)
    date_col_vals = ws.col_values(REPORT_DATE_COL)  # значения столбца B сверху вниз
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

    # Готовим точечные обновления ячеек
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
    # колонки фракций: индексы 5..9 (0-based) при PROD_HEADERS
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
    # Реализация: Фракция = индекс 3, Количество кг = индекс 4
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
{"дата":"дд.мм.гггг","фио":"ФИО","вес_шин":0,"мешки":0,"нитки":0,"фракция_0_1":0,"фракция_1_2":0,"фракция_2_4":0,"фракция_4_6":0,"фракция_6_8":0,"металл_корд":0,"примечание":""}
Если не читается — 0."""
    return await call_claude(image_b64, prompt)


async def recognize_sale(image_bytes: bytes):
    image_b64 = base64.b64encode(image_bytes).decode()
    # CHANGED: добавлено поле цена_за_кг
    prompt = """Это фото накладной на отпуск/реализацию резиновой крошки.
Распознай данные и верни ТОЛЬКО JSON без markdown:
{"дата":"дд.мм.гггг","покупатель":"название организации или ИП","фракция":"например 2-4","количество_т":0.0,"количество_кг":0,"цена_за_кг":0,"сумма_с_ндс":0,"сумма_ндс":0,"примечание":""}
Количество бери из колонки "отпущено". Если оно в тоннах — заполни количество_т, если в кг — количество_кг.
Цену бери из колонки "цена" (цена за 1 кг). Бери ровно как в накладной, не пересчитывай.
Сумму с НДС бери из колонки "Сумма с НДС" (полная сумма включая НДС).
Сумму НДС бери из колонки "Сумма НДС" (только НДС). Если НДС в накладной нет — поставь 0.
Если не читается — 0."""
    return await call_claude(image_b64, prompt)


# ---------- Команды и меню ----------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("⛔ Нет доступа. Обратитесь к администратору.")
        return
    keyboard = [
        ["📸 Производство", "📄 Реализация"],
        ["✍️ Ввод производства", "✍️ Ввод реализации"],
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
    data_rows = rows[1:]  # без шапки
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
        t = datetime.now().strftime("%d.%m.%Y")
    context.user_data["prod"]["дата"] = t
    await update.message.reply_text("👤 ФИО оператора?")
    return P_FIO


async def manual_prod_fio(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["prod"]["фио"] = update.message.text.strip()
    await update.message.reply_text("⚖️ Вес шин, кг?")
    return P_TIRES


async def manual_prod_tires(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["prod"]["вес_шин"] = parse_num(update.message.text)
    await update.message.reply_text("🛍 Мешки, шт?")
    return P_BAGS


async def manual_prod_bags(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["prod"]["мешки"] = parse_num(update.message.text)
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


async def manual_prod_note(update: Update, context: ContextTypes.DEFAULT_TYPE):
    note = update.message.text.strip()
    context.user_data["prod"]["примечание"] = "" if note in ("-", "—") else note
    d = context.user_data["prod"]

    existing = get_existing_dates_production()
    dup_note = ""
    if str(d.get("дата", "")).strip() in existing:
        dup_note = "\n\n⚠️ Запись с этой датой уже есть — добавлю ещё одну строку."

    save_production(d)
    total = (parse_num(d.get("фракция_0_1", 0)) + parse_num(d.get("фракция_1_2", 0))
             + parse_num(d.get("фракция_2_4", 0)) + parse_num(d.get("фракция_4_6", 0))
             + parse_num(d.get("фракция_6_8", 0)))
    reply = (
        f"✅ *Производство сохранено!*\n\n"
        f"📅 Дата: {d.get('дата', '—')}\n"
        f"👤 Оператор: {d.get('фио', '—')}\n"
        f"⚖️ Вес шин: {d.get('вес_шин', 0)} кг\n"
        f"🛍 Мешки: {d.get('мешки', 0)} шт\n\n"
        f"*Фракции (кг):*\n"
        f"  0-1: {d.get('фракция_0_1', 0)}\n"
        f"  1-2: {d.get('фракция_1_2', 0)}\n"
        f"  2-4: {d.get('фракция_2_4', 0)}\n"
        f"  4-6: {d.get('фракция_4_6', 0)}\n"
        f"  6-8: {d.get('фракция_6_8', 0)}\n"
        f"  Всего: {total} кг\n\n"
        f"🔩 Металлокорд: {d.get('металл_корд', 0)} кг" + dup_note
    )
    await update.message.reply_text(reply, parse_mode="Markdown")
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
        t = datetime.now().strftime("%d.%m.%Y")
    context.user_data["sale"]["дата"] = t
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
    # CHANGED: теперь спрашиваем цену за кг
    await update.message.reply_text("💵 Цена за кг, тнг? (как в накладной)")
    return S_PRICE


async def manual_sale_price(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # NEW: шаг "цена за кг"
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

    # NEW: дописываем в отчёт; ошибка отчёта не валит сохранение в склад
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
    await update.message.reply_text("❌ Ручной ввод отменён.")
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


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("⛔ Нет доступа. Обратитесь к администратору.")
        return
    photo_type = context.user_data.get("photo_type", "production")
    await update.message.reply_text("📸 Получил фото, распознаю данные...")

    try:
        photo = update.message.photo[-1]
        file = await context.bot.get_file(photo.file_id)
        image_bytes = bytes(await file.download_as_bytearray())

        if photo_type == "sale":
            data = await recognize_sale(image_bytes)
            logger.info(f"Sale data: {data}")
            # Тип расчёта с фото не всегда виден — по умолчанию пусто, уточнит вручную при необходимости
            data.setdefault("тип_расчета", "")
            save_sale(data)

            # NEW: пробуем дописать в отчёт (только если тип расчёта распознан)
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
            context.user_data["photo_type"] = "production"

        else:
            data = await recognize_production(image_bytes)
            logger.info(f"Production data: {data}")
            existing_dates = get_existing_dates_production()
            records = data if isinstance(data, list) else [data]
            new_records = [r for r in records if str(r.get("дата", "")).strip() not in existing_dates]
            skipped = len(records) - len(new_records)

            if not new_records:
                await update.message.reply_text("⚠️ Все записи с этого фото уже есть в таблице.")
                return

            for item in new_records:
                save_production(item)

            first = new_records[-1]
            reply = (
                f"✅ *Производство сохранено!*\n"
                f"_Новых: {len(new_records)}"
                + (f" | Дублей пропущено: {skipped}" if skipped > 0 else "") +
                "_\n\n"
                f"📅 Дата: {first.get('дата', '—')}\n"
                f"👤 Оператор: {first.get('фио', '—')}\n"
                f"⚖️ Вес шин: {first.get('вес_шин', 0)} кг\n\n"
                f"*Фракции (кг):*\n"
                f"  0-1: {first.get('фракция_0_1', 0)}\n"
                f"  1-2: {first.get('фракция_1_2', 0)}\n"
                f"  2-4: {first.get('фракция_2_4', 0)}\n"
                f"  4-6: {first.get('фракция_4_6', 0)}\n"
                f"  6-8: {first.get('фракция_6_8', 0)}\n\n"
                f"🔩 Металлокорд: {first.get('металл_корд', 0)} кг"
            )
            await update.message.reply_text(reply, parse_mode="Markdown")

    except Exception as e:
        logger.error(f"Error: {e}")
        await update.message.reply_text(
            "❌ Не смог распознать фото. Попробуй:\n"
            "• Сделать фото чётче\n"
            "• Хорошее освещение\n"
            "• Всё в кадре"
        )


def main():
    app = Application.builder().token(BOT_TOKEN).build()

    prod_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^✍️ Ввод производства$"), manual_prod_start)],
        states={
            P_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_date)],
            P_FIO: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_fio)],
            P_TIRES: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_tires)],
            P_BAGS: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_bags)],
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
            S_PRICE: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_price)],  # NEW
            S_SUM_VAT: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_sum_vat)],
            S_VAT: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_vat)],
            S_NOTE: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_note)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("ostatok", ostatok))
    app.add_handler(CommandHandler("last", last_records))
    app.add_handler(CommandHandler("cancel", cancel))
    app.add_handler(prod_conv)
    app.add_handler(sale_conv)
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
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
