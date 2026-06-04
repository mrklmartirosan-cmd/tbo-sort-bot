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
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
WEBHOOK_URL = os.environ.get("WEBHOOK_URL", "")
PORT = int(os.environ.get("PORT", 8080))

PRODUCTION_FILE = "uchet_kroshki.xlsx"
SALES_FILE = "uchet_realizacii.xlsx"

# Состояния диалогов ручного ввода
(P_DATE, P_FIO, P_TIRES, P_F01, P_F12, P_F24, P_F46, P_F68, P_CORD, P_NOTE) = range(10)
(S_DATE, S_BUYER, S_FRAC, S_KG, S_SUM_VAT, S_VAT, S_NOTE) = range(10, 17)

FRACTIONS = ["0-1", "1-2", "2-4", "4-6", "6-8"]


def make_border():
    thin = Side(style='thin', color='AAAAAA')
    return Border(left=thin, right=thin, top=thin, bottom=thin)


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


def get_or_create_production():
    if os.path.exists(PRODUCTION_FILE):
        return load_workbook(PRODUCTION_FILE)
    wb = Workbook()
    ws = wb.active
    ws.title = "Журнал"
    border = make_border()
    headers = ["Дата", "ФИО оператора", "Вес шин, кг", "Фракция 0-1", "Фракция 1-2",
               "Фракция 2-4", "Фракция 4-6", "Фракция 6-8", "Всего крошки, кг", "Металл. корд, кг", "Примечание"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col)
        c.value = h
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor="1F5C8B")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws.row_dimensions[1].height = 36
    for i, w in enumerate([12, 20, 12, 11, 11, 11, 11, 11, 14, 13, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    wb.save(PRODUCTION_FILE)
    return wb


def get_or_create_sales():
    if os.path.exists(SALES_FILE):
        return load_workbook(SALES_FILE)
    wb = Workbook()
    ws = wb.active
    ws.title = "Реализация"
    border = make_border()
    headers = ["Дата", "Покупатель", "Фракция", "Количество, т", "Количество, кг",
               "Сумма с НДС, тнг", "Сумма НДС, тнг", "Примечание"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col)
        c.value = h
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor="1E7145")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws.row_dimensions[1].height = 36
    for i, w in enumerate([12, 25, 15, 14, 14, 18, 16, 20], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    wb.save(SALES_FILE)
    return wb


def get_existing_dates_production():
    wb = get_or_create_production()
    ws = wb.active
    dates = set()
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row=row, column=1).value
        if val:
            dates.add(str(val).strip())
    return dates


def save_production(data):
    wb = get_or_create_production()
    ws = wb.active
    border = make_border()
    next_row = ws.max_row + 1
    shade = "F2F2F2" if next_row % 2 == 0 else "FFFFFF"
    values = [
        data.get("дата", datetime.now().strftime("%d.%m.%Y")),
        data.get("фио", ""),
        data.get("вес_шин", ""),
        data.get("фракция_0_1", ""),
        data.get("фракция_1_2", ""),
        data.get("фракция_2_4", ""),
        data.get("фракция_4_6", ""),
        data.get("фракция_6_8", ""),
        f"=SUM(D{next_row}:H{next_row})",
        data.get("металл_корд", ""),
        data.get("примечание", ""),
    ]
    for col, val in enumerate(values, 1):
        c = ws.cell(row=next_row, column=col)
        c.value = val
        c.font = Font(name="Arial", size=10)
        c.fill = PatternFill("solid", fgColor=shade)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    wb.save(PRODUCTION_FILE)


def save_sale(data):
    wb = get_or_create_sales()
    ws = wb.active
    border = make_border()
    next_row = ws.max_row + 1
    shade = "F2F2F2" if next_row % 2 == 0 else "FFFFFF"
    # Источник правды — кг. Тонны считаем из кг для отображения.
    qty_kg = parse_num(data.get("количество_кг", 0))
    if not qty_kg and data.get("количество_т"):
        qty_kg = parse_num(data.get("количество_т", 0)) * 1000
    qty_t = round(qty_kg / 1000, 3) if qty_kg else 0
    values = [
        data.get("дата", datetime.now().strftime("%d.%m.%Y")),
        data.get("покупатель", ""),
        data.get("фракция", ""),
        qty_t,
        qty_kg,
        parse_num(data.get("сумма_с_ндс", 0)),
        parse_num(data.get("сумма_ндс", 0)),
        data.get("примечание", ""),
    ]
    for col, val in enumerate(values, 1):
        c = ws.cell(row=next_row, column=col)
        c.value = val
        c.font = Font(name="Arial", size=10)
        c.fill = PatternFill("solid", fgColor=shade)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    wb.save(SALES_FILE)


def calc_stock():
    income = {k: 0 for k in FRACTIONS}
    wb_p = get_or_create_production()
    ws_p = wb_p.active
    for row in range(2, ws_p.max_row + 1):
        for i, key in enumerate(FRACTIONS, 4):
            val = ws_p.cell(row=row, column=i).value
            try:
                income[key] += float(val) if val else 0
            except Exception:
                pass

    outcome = {k: 0 for k in FRACTIONS}
    wb_s = get_or_create_sales()
    ws_s = wb_s.active
    for row in range(2, ws_s.max_row + 1):
        frac = ws_s.cell(row=row, column=3).value or ""
        qty_kg = ws_s.cell(row=row, column=5).value or 0
        for key in outcome:
            if key in str(frac):
                try:
                    outcome[key] += float(qty_kg)
                except Exception:
                    pass
                break

    stock = {k: income[k] - outcome[k] for k in income}
    return income, outcome, stock


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
{"дата":"дд.мм.гггг","фио":"ФИО","вес_шин":0,"фракция_0_1":0,"фракция_1_2":0,"фракция_2_4":0,"фракция_4_6":0,"фракция_6_8":0,"металл_корд":0,"примечание":""}
Если не читается — 0."""
    return await call_claude(image_b64, prompt)


async def recognize_sale(image_bytes: bytes):
    image_b64 = base64.b64encode(image_bytes).decode()
    prompt = """Это фото накладной на отпуск/реализацию резиновой крошки.
Распознай данные и верни ТОЛЬКО JSON без markdown:
{"дата":"дд.мм.гггг","покупатель":"название организации или ИП","фракция":"например 2-4","количество_т":0.0,"количество_кг":0,"сумма_с_ндс":0,"сумма_ндс":0,"примечание":""}
Количество бери из колонки "отпущено". Если оно в тоннах — заполни количество_т, если в кг — количество_кг.
Сумму с НДС бери из колонки "Сумма с НДС" (полная сумма включая НДС).
Сумму НДС бери из колонки "Сумма НДС" (только НДС).
Если не читается — 0."""
    return await call_claude(image_b64, prompt)


# ---------- Команды и меню ----------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
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
        "/get — скачать таблицы Excel\n"
        "/last — последние записи\n"
        "/cancel — отменить ручной ввод",
        reply_markup=reply_markup
    )


async def ostatok(update: Update, context: ContextTypes.DEFAULT_TYPE):
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


async def get_files(update: Update, context: ContextTypes.DEFAULT_TYPE):
    date_str = datetime.now().strftime('%d%m%Y')
    get_or_create_production()
    get_or_create_sales()
    for filepath, caption in [
        (PRODUCTION_FILE, "📊 Журнал производства"),
        (SALES_FILE, "📄 Журнал реализации")
    ]:
        with open(filepath, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=f"{filepath.replace('.xlsx', '')}_{date_str}.xlsx",
                caption=caption
            )


async def last_records(update: Update, context: ContextTypes.DEFAULT_TYPE):
    wb = get_or_create_production()
    ws = wb.active
    max_row = ws.max_row
    if max_row <= 1:
        await update.message.reply_text("📭 Записей пока нет.")
        return
    start_row = max(2, max_row - 4)
    text = "📋 *Последние записи производства:*\n\n"
    for row in range(start_row, max_row + 1):
        date = ws.cell(row=row, column=1).value or "—"
        fio = ws.cell(row=row, column=2).value or "—"
        total = ws.cell(row=row, column=9).value or "—"
        text += f"📅 {date} | {fio} | {total} кг\n"
    await update.message.reply_text(text, parse_mode="Markdown")


# ---------- Ручной ввод: ПРОИЗВОДСТВО ----------

async def manual_prod_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
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
    total = (d.get("фракция_0_1", 0) + d.get("фракция_1_2", 0) + d.get("фракция_2_4", 0)
             + d.get("фракция_4_6", 0) + d.get("фракция_6_8", 0))
    reply = (
        f"✅ *Производство сохранено!*\n\n"
        f"📅 Дата: {d.get('дата', '—')}\n"
        f"👤 Оператор: {d.get('фио', '—')}\n"
        f"⚖️ Вес шин: {d.get('вес_шин', 0)} кг\n\n"
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
    qty_kg = parse_num(d.get("количество_кг", 0))
    reply = (
        f"✅ *Реализация сохранена!*\n\n"
        f"📅 Дата: {d.get('дата', '—')}\n"
        f"🏢 Покупатель: {d.get('покупатель', '—')}\n"
        f"📦 Фракция: {d.get('фракция', '—')}\n"
        f"⚖️ Количество: {qty_kg} кг\n"
        f"💰 Сумма с НДС: {d.get('сумма_с_ндс', 0)} тнг\n"
        f"💰 Сумма НДС: {d.get('сумма_ндс', 0)} тнг"
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
    text = update.message.text
    if text == "📸 Производство":
        context.user_data["photo_type"] = "production"
        await update.message.reply_text("📸 Отправь фото отчёта по производству крошки.")
    elif text == "📄 Реализация":
        context.user_data["photo_type"] = "sale"
        await update.message.reply_text("📄 Отправь фото накладной на реализацию.")
    else:
        await update.message.reply_text("Используй кнопки меню или команды /ostatok, /get, /last")


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    photo_type = context.user_data.get("photo_type", "production")
    await update.message.reply_text("📸 Получил фото, распознаю данные...")

    try:
        photo = update.message.photo[-1]
        file = await context.bot.get_file(photo.file_id)
        image_bytes = bytes(await file.download_as_bytearray())

        if photo_type == "sale":
            data = await recognize_sale(image_bytes)
            logger.info(f"Sale data: {data}")
            save_sale(data)
            qty_kg = parse_num(data.get("количество_кг", 0))
            if not qty_kg and data.get("количество_т"):
                qty_kg = parse_num(data.get("количество_т", 0)) * 1000
            reply = (
                f"✅ *Реализация сохранена!*\n\n"
                f"📅 Дата: {data.get('дата', '—')}\n"
                f"🏢 Покупатель: {data.get('покупатель', '—')}\n"
                f"📦 Фракция: {data.get('фракция', '—')}\n"
                f"⚖️ Количество: {qty_kg:.0f} кг\n"
                f"💰 Сумма с НДС: {data.get('сумма_с_ндс', 0)} тнг\n"
                f"💰 Сумма НДС: {data.get('сумма_ндс', 0)} тнг"
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

    # Диалог ручного ввода производства
    prod_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^✍️ Ввод производства$"), manual_prod_start)],
        states={
            P_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_date)],
            P_FIO: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_fio)],
            P_TIRES: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prod_tires)],
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

    # Диалог ручного ввода реализации
    sale_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^✍️ Ввод реализации$"), manual_sale_start)],
        states={
            S_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_date)],
            S_BUYER: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_buyer)],
            S_FRAC: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_frac)],
            S_KG: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_kg)],
            S_SUM_VAT: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_sum_vat)],
            S_VAT: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_vat)],
            S_NOTE: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_sale_note)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("ostatok", ostatok))
    app.add_handler(CommandHandler("get", get_files))
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
