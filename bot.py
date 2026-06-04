import os
import json
import base64
import logging
import re
from datetime import datetime
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import httpx
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN", "8301597645:AAH1YI80SUG0439UJTHqyw8jhsPfNydgWrg")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "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")
WEBHOOK_URL = os.environ.get("WEBHOOK_URL", "")
PORT = int(os.environ.get("PORT", 8080))
EXCEL_FILE = "uchet_kroshki_bot.xlsx"

def get_or_create_excel():
    if os.path.exists(EXCEL_FILE):
        return load_workbook(EXCEL_FILE)
    wb = Workbook()
    ws = wb.active
    ws.title = "Журнал"
    thin = Side(style='thin', color='AAAAAA')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
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
    col_widths = [12, 20, 12, 11, 11, 11, 11, 11, 14, 13, 18]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    wb.save(EXCEL_FILE)
    return wb

def save_to_excel(data: dict):
    wb = get_or_create_excel()
    ws = wb.active
    thin = Side(style='thin', color='AAAAAA')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
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
    wb.save(EXCEL_FILE)

async def recognize_photo(image_bytes: bytes) -> dict:
    image_b64 = base64.b64encode(image_bytes).decode()
    prompt = """Это фото ежедневного отчёта по производству резиновой крошки.
Распознай все данные и верни ТОЛЬКО JSON без пояснений и без markdown:
{
  "дата": "дд.мм.гггг",
  "фио": "ФИО оператора",
  "вес_шин": 0,
  "фракция_0_1": 0,
  "фракция_1_2": 0,
  "фракция_2_4": 0,
  "фракция_4_6": 0,
  "фракция_6_8": 0,
  "металл_корд": 0,
  "примечание": ""
}
Если значение не читается — поставь 0."""
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
                "max_tokens": 1000,
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
        text = result["content"][0]["text"].strip()
        text = re.sub(r'```json|```', '', text).strip()
        return json.loads(text)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Привет! Я бот учёта производства резиновой крошки.\n\n"
        "📸 Отправь фото ежедневного отчёта — распознаю и сохраню в таблицу.\n\n"
        "📊 Команды:\n"
        "/get — скачать Excel таблицу\n"
        "/last — последние 5 записей"
    )

async def get_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    get_or_create_excel()
    with open(EXCEL_FILE, "rb") as f:
        await update.message.reply_document(
            document=f,
            filename=f"uchet_kroshki_{datetime.now().strftime('%d%m%Y')}.xlsx",
            caption="📊 Текущая таблица учёта"
        )

async def last_records(update: Update, context: ContextTypes.DEFAULT_TYPE):
    wb = get_or_create_excel()
    ws = wb.active
    max_row = ws.max_row
    if max_row <= 1:
        await update.message.reply_text("📭 Записей пока нет.")
        return
    start_row = max(2, max_row - 4)
    text = "📋 *Последние записи:*\n\n"
    for row in range(start_row, max_row + 1):
        date = ws.cell(row=row, column=1).value or "—"
        fio = ws.cell(row=row, column=2).value or "—"
        total = ws.cell(row=row, column=9).value or "—"
        text += f"📅 {date} | {fio} | Крошка: {total} кг\n"
    await update.message.reply_text(text, parse_mode="Markdown")

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📸 Получил фото, распознаю данные...")
    try:
        photo = update.message.photo[-1]
        file = await context.bot.get_file(photo.file_id)
        image_bytes = await file.download_as_bytearray()
        data = await recognize_photo(bytes(image_bytes))
        save_to_excel(data)
        reply = (
            f"✅ *Данные распознаны и сохранены!*\n\n"
            f"📅 Дата: {data.get('дата', '—')}\n"
            f"👤 Оператор: {data.get('фио', '—')}\n"
            f"⚖️ Вес шин: {data.get('вес_шин', 0)} кг\n\n"
            f"*Фракции (кг):*\n"
            f"  0-1: {data.get('фракция_0_1', 0)}\n"
            f"  1-2: {data.get('фракция_1_2', 0)}\n"
            f"  2-4: {data.get('фракция_2_4', 0)}\n"
            f"  4-6: {data.get('фракция_4_6', 0)}\n"
            f"  6-8: {data.get('фракция_6_8', 0)}\n\n"
            f"🔩 Металлокорд: {data.get('металл_корд', 0)} кг\n\n"
            f"_Если что-то неверно — напиши мне_"
        )
        await update.message.reply_text(reply, parse_mode="Markdown")
    except Exception as e:
        logger.error(f"Error: {e}")
        await update.message.reply_text(
            "❌ Не смог распознать. Попробуй:\n"
            "• Сделать фото чётче\n"
            "• Хорошее освещение\n"
            "• Всё в кадре"
        )

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📸 Отправь фото отчёта или используй /get для скачивания таблицы.")

def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("get", get_excel))
    app.add_handler(CommandHandler("last", last_records))
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
