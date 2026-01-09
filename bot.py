# bot.py — версия для Render.com
import os
import tempfile
import pandas as pd
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, filters, ContextTypes
)
from telegram.constants import ParseMode
import logging

# Отключаем лишние логи
logging.basicConfig(level=logging.INFO)
logging.getLogger("httpx").setLevel(logging.WARNING)

# === НАСТРОЙКА ===
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
if not TELEGRAM_TOKEN:
    raise ValueError("❌ Переменная окружения TELEGRAM_TOKEN не задана!")

user_files = {}

def parse_time_safe(val):
    if pd.isna(val):
        return None
    if isinstance(val, str):
        for fmt in ["%H:%M:%S", "%H:%M", "%H:%M:%S.%f"]:
            try:
                t = datetime.strptime(val, fmt).time()
                return datetime.combine(datetime.today(), t)
            except:
                continue
    elif hasattr(val, 'time'):
        return val
    return None

def generate_report(df, report_type="full"):
    exclude = ["(без ответственного)", "IT Отдел"]
    df = df[~df["Сотрудник"].isin(exclude)].copy()
    df["Время_dt"] = df["Время"].apply(parse_time_safe)
    df = df[df["Время_dt"].notna()].copy()

    if df.empty:
        return "Нет данных после фильтрации."

    df["Минута"] = df["Время_dt"].dt.floor('min')
    df["Минута_str"] = df["Минута"].dt.strftime("%H:%M")
    unique_minutes = sorted(df["Минута_str"].unique())[-10:]

    lines = []
    for minute in unique_minutes:
        group = df[df["Минута_str"] == minute]
        total = len(group)
        if total == 0:
            continue

        ao = group["Результат"].str.contains(
            r"Автоответчик|Обнаружен автоответчик \(системный\)",
            na=False, regex=True
        ).sum()
        silence = (group["Результат"] == "Тишина").sum()
        managers = group["Сотрудник"].nunique()

        avg = round(total / managers, 2) if managers > 0 else 0.0
        pct_ao = round(ao / total * 100, 2) if total > 0 else 0.0
        pct_silence = round(silence / total * 100, 2) if total > 0 else 0.0

        if report_type == "full":
            lines.append(
                f"Статистика с {minute} по {minute}\n"
                f"- Звонков: {total}\n"
                f"- АО: {ao}\n"
                f"- Тишина: {silence}\n"
                f"- <b>Процент АО: {pct_ao}%</b>\n"
                f"- Процент Тишина: {pct_silence}%\n"
                f"- В среднем звонков: {avg}\n"
            )
        elif report_type == "ao_only":
            lines.append(
                f"Статистика с {minute} по {minute}\n"
                f"- Звонков: {total}\n"
                f"- <b>Процент АО: {pct_ao}%</b>\n"
            )
        elif report_type == "silence_only":
            lines.append(
                f"Статистика с {minute} по {minute}\n"
                f"- Звонков: {total}\n"
                f"- Процент Тишина: {pct_silence}%\n"
            )
        elif report_type == "ao_silence":
            lines.append(
                f"Статистика с {minute} по {minute}\n"
                f"- Звонков: {total}\n"
                f"- <b>Процент АО: {pct_ao}%</b>\n"
                f"- Процент Тишина: {pct_silence}%\n"
            )

    return "\n".join(lines).strip() or "Нет данных за последние 10 минут."

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Отправьте Excel-файл с отчётом Скорозвона.")

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.endswith(('.xlsx', '.xls')):
        await update.message.reply_text("Пожалуйста, отправьте файл в формате Excel (.xlsx).")
        return

    try:
        file = await doc.get_file()
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            await file.download_to_drive(tmp.name)
            user_files[update.effective_user.id] = tmp.name

        keyboard = [
            [InlineKeyboardButton("Полный", callback_data="full"),
             InlineKeyboardButton("По % АО", callback_data="ao_only")],
            [InlineKeyboardButton("По % Тишины", callback_data="silence_only"),
             InlineKeyboardButton("АО + Тишина", callback_data="ao_silence")]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await update.message.reply_text("Выберите тип отчёта:", reply_markup=reply_markup)

    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка загрузки файла:\n{str(e)}")

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = update.effective_user.id
    report_type = query.data

    file_path = user_files.get(user_id)
    if not file_path or not os.path.exists(file_path):
        await query.edit_message_text("Файл утерян. Отправьте Excel заново.")
        return

    try:
        df = pd.read_excel(file_path)

        # Если колонки Unnamed → значит, нет заголовков
        if all(str(col).startswith("Unnamed") for col in df.columns[:3]):
            df = pd.read_excel(file_path, header=None)
            if df.shape[1] < 3:
                raise ValueError("Файл содержит меньше 3 колонок")
            df.columns = ["Время", "Результат", "Сотрудник"] + [f"col_{i}" for i in range(3, df.shape[1])]

        required = ["Время", "Результат", "Сотрудник"]
        missing = [col for col in required if col not in df.columns]
        if missing:
            raise ValueError(f"Не найдены колонки: {missing}")

        report = generate_report(df, report_type)

        # Удаляем файл
        try:
            os.unlink(file_path)
        except:
            pass
        user_files.pop(user_id, None)

        # Отправка с разбивкой
        MAX_LEN = 4000
        if len(report) <= MAX_LEN:
            await query.edit_message_text(report, parse_mode=ParseMode.HTML)
        else:
            await query.edit_message_text("Отчёт большой. Отправляю частями...")
            parts = []
            temp = report
            while temp:
                if len(temp) <= MAX_LEN:
                    parts.append(temp)
                    break
                split_idx = temp[:MAX_LEN].rfind('\n\n')
                if split_idx == -1:
                    split_idx = MAX_LEN
                parts.append(temp[:split_idx])
                temp = temp[split_idx:].lstrip()

            for i, part in enumerate(parts):
                if i == 0:
                    await query.edit_message_text(part, parse_mode=ParseMode.HTML)
                else:
                    await context.bot.send_message(
                        chat_id=query.message.chat_id,
                        text=part,
                        parse_mode=ParseMode.HTML
                    )

    except Exception as e:
        error_msg = f"❌ Ошибка анализа:\n{str(e)}"
        if len(error_msg) > 4096:
            error_msg = error_msg[:4093] + "..."
        await query.edit_message_text(error_msg)

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(CallbackQueryHandler(button_handler))
    print("✅ Бот запущен на Render!")
    app.run_polling()

if __name__ == "__main__":
    main()