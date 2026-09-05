
import os
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from datetime import datetime, date, timedelta

import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    ConversationHandler, ContextTypes, filters
)

TOKEN = os.environ["BOT_TOKEN"]
SHEET_ID = os.environ["GOOGLE_SHEET_ID"]
WEBHOOK_URL = os.environ["WEBHOOK_URL"].rstrip("/")
WEBHOOK_PATH = os.environ.get("WEBHOOK_PATH", "telegram")
PORT = int(os.environ.get("PORT", "10000"))

CREDENTIALS_FILE = os.environ.get(
    "GOOGLE_APPLICATION_CREDENTIALS",
    "/etc/secrets/google-service-account.json"
)

TRACTORS = [
    "60059KBA","60227SBA","60510KBA","60501KBA","60507KBA","60511KBA",
    "60260SBA","60515KBA","60277KBA","60110QBA","60720SBA","60545SBA",
    "60740SBA","60730SBA","60474SBA","60373SBA","60414SBA","60755SBA",
    "60441SBA","60442SBA","60445SBA"
]

MENU = ReplyKeyboardMarkup(
    [
        ["➕ Navbat berildi", "💰 Shoferdan olindi"],
        ["📊 Hisobot", "📥 Excel"],
    ],
    resize_keyboard=True
)

ASK_DATE, ASK_TRUCK, ASK_AMOUNT = range(3)

# Google Sheets connection
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
gc = gspread.authorize(creds)
sh = gc.open_by_key(SHEET_ID)

try:
    ws = sh.worksheet("Navbat")
except gspread.WorksheetNotFound:
    ws = sh.add_worksheet(title="Navbat", rows=1000, cols=10)

HEADERS = [
    "ID", "Holat", "Berilgan sana", "Mashina",
    "Berilgan summa", "Olingan sana", "Olingan summa",
    "Farq", "Yaratilgan vaqt"
]

def ensure_headers():
    first = ws.row_values(1)
    if first != HEADERS:
        ws.update("A1:I1", [HEADERS])

ensure_headers()

def rows():
    values = ws.get_all_values()
    if len(values) <= 1:
        return []
    result = []
    for r in values[1:]:
        r = r + [""] * (9 - len(r))
        result.append(r[:9])
    return result

def next_id():
    vals = rows()
    nums = []
    for r in vals:
        try:
            nums.append(int(r[0]))
        except Exception:
            pass
    return str(max(nums, default=0) + 1)

def add_given(truck, dt, amount):
    row = [
        next_id(), "BERILGAN", dt, truck,
        float(amount), "", "", float(amount), datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ]
    ws.append_row(row, value_input_option="USER_ENTERED")

def receive_latest(truck, dt, amount):
    values = ws.get_all_values()
    target_row = None
    target_data = None

    for i in range(len(values) - 1, 0, -1):
        r = values[i] + [""] * (9 - len(values[i]))
        if r[3] == truck and r[1] == "BERILGAN":
            target_row = i + 1
            target_data = r[:9]
            break

    if target_row is None:
        return False

    given = float(target_data[4] or 0)
    received = float(amount)
    diff = given - received

    ws.update(
        f"B{target_row}:H{target_row}",
        [["YOPILDI", target_data[2], target_data[3], given, dt, received, diff]]
    )
    return True

def money(x):
    try:
        return f"{float(x):,.0f}".replace(",", " ")
    except Exception:
        return "0"

def report_text():
    data = rows()
    totals = {t: [0.0, 0.0] for t in TRACTORS}

    for r in data:
        truck = r[3]
        if truck not in totals:
            totals[truck] = [0.0, 0.0]
        try:
            totals[truck][0] += float(r[4] or 0)
        except Exception:
            pass
        try:
            totals[truck][1] += float(r[6] or 0)
        except Exception:
            pass

    lines = ["📊 WikCar NAVBAT HISOBOTI", ""]
    total_given = total_received = 0.0

    for truck in TRACTORS:
        given, received = totals[truck]
        diff = given - received
        total_given += given
        total_received += received
        lines.append(
            f"{truck}: Berilgan {money(given)} | Olingan {money(received)} | Farq {money(diff)}"
        )

    lines += [
        "",
        "━━━━━━━━━━━━━━━━━━",
        f"Jami berilgan: {money(total_given)}",
        f"Jami olingan: {money(total_received)}",
        f"Jami farq: {money(total_given - total_received)}",
    ]
    return "\n".join(lines)

def make_excel(path="/tmp/navbat.xlsx"):
    data = rows()
    wb = Workbook()
    wsx = wb.active
    wsx.title = "Navbat"
    wsx.append(HEADERS)
    for r in data:
        wsx.append(r)

    report = wb.create_sheet("Hisobot")
    report.append(["Mashina", "Berilgan summa", "Olingan summa", "Farq"])

    totals = {t: [0.0, 0.0] for t in TRACTORS}
    for r in data:
        truck = r[3]
        if truck not in totals:
            totals[truck] = [0.0, 0.0]
        try: totals[truck][0] += float(r[4] or 0)
        except: pass
        try: totals[truck][1] += float(r[6] or 0)
        except: pass

    for truck in TRACTORS:
        g, o = totals[truck]
        report.append([truck, g, o, g-o])

    tg = sum(v[0] for v in totals.values())
    to = sum(v[1] for v in totals.values())
    report.append(["JAMI", tg, to, tg-to])
    wb.save(path)
    return path

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Assalomu alaykum! WikCar Navbat bot.", reply_markup=MENU)

async def menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if text == "📊 Hisobot":
        await update.message.reply_text(report_text(), reply_markup=MENU)
        return ConversationHandler.END

    if text == "📥 Excel":
        path = make_excel()
        with open(path, "rb") as f:
            await update.message.reply_document(f, filename="WikCar_Navbat.xlsx", reply_markup=MENU)
        return ConversationHandler.END

    if text == "➕ Navbat berildi":
        context.user_data["mode"] = "given"
        kb = [
            [InlineKeyboardButton("Bugun", callback_data="date:today"),
             InlineKeyboardButton("Kecha", callback_data="date:yesterday")],
            [InlineKeyboardButton("Boshqa sana", callback_data="date:other")]
        ]
        await update.message.reply_text("Sanani tanlang:", reply_markup=InlineKeyboardMarkup(kb))
        return ASK_DATE

    if text == "💰 Shoferdan olindi":
        context.user_data["mode"] = "received"
        kb = [
            [InlineKeyboardButton("Bugun", callback_data="date:today"),
             InlineKeyboardButton("Kecha", callback_data="date:yesterday")],
            [InlineKeyboardButton("Boshqa sana", callback_data="date:other")]
        ]
        await update.message.reply_text("Sanani tanlang:", reply_markup=InlineKeyboardMarkup(kb))
        return ASK_DATE

async def date_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    value = q.data.split(":")[1]

    if value == "today":
        context.user_data["date"] = date.today().isoformat()
    elif value == "yesterday":
        context.user_data["date"] = (date.today() - timedelta(days=1)).isoformat()
    else:
        await q.edit_message_text("Sanani YYYY-MM-DD ko‘rinishida yuboring:")
        context.user_data["waiting_custom_date"] = True
        return ASK_DATE

    return await show_trucks(q, context)

async def custom_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.user_data.get("waiting_custom_date"):
        return ASK_DATE
    try:
        d = datetime.strptime(update.message.text.strip(), "%Y-%m-%d").date()
        context.user_data["date"] = d.isoformat()
        context.user_data["waiting_custom_date"] = False
        return await show_trucks(update, context)
    except ValueError:
        await update.message.reply_text("Sana noto‘g‘ri. Masalan: 2026-09-05")
        return ASK_DATE

async def show_trucks(target, context):
    buttons = []
    for i in range(0, len(TRACTORS), 3):
        buttons.append([
            InlineKeyboardButton(t, callback_data=f"truck:{t}")
            for t in TRACTORS[i:i+3]
        ])
    markup = InlineKeyboardMarkup(buttons)
    if hasattr(target, "edit_message_text"):
        await target.edit_message_text("Mashina tanlang:", reply_markup=markup)
    else:
        await target.message.reply_text("Mashina tanlang:", reply_markup=markup)
    return ASK_TRUCK

async def truck_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    truck = q.data.split(":", 1)[1]
    context.user_data["truck"] = truck
    await q.edit_message_text(
        f"Mashina: {truck}\nSummani kiriting:"
    )
    return ASK_AMOUNT

async def amount_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip().replace(" ", "").replace(",", ".")
    try:
        amount = float(text)
        if amount < 0:
            raise ValueError
    except ValueError:
        await update.message.reply_text("Summani raqamda kiriting. Masalan: 150000")
        return ASK_AMOUNT

    truck = context.user_data["truck"]
    dt = context.user_data["date"]
    mode = context.user_data["mode"]

    if mode == "given":
        add_given(truck, dt, amount)
        await update.message.reply_text(
            f"✅ Navbat berildi\n🚛 {truck}\n📅 {dt}\n💰 {money(amount)}",
            reply_markup=MENU
        )
    else:
        ok = receive_latest(truck, dt, amount)
        if ok:
            await update.message.reply_text(
                f"✅ Shoferdan olindi\n🚛 {truck}\n📅 {dt}\n💰 {money(amount)}",
                reply_markup=MENU
            )
        else:
            await update.message.reply_text(
                f"❌ {truck} uchun ochiq navbat topilmadi.",
                reply_markup=MENU
            )

    context.user_data.clear()
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("Bekor qilindi.", reply_markup=MENU)
    return ConversationHandler.END

def main():
    app = Application.builder().token(TOKEN).build()

    conv = ConversationHandler(
        entry_points=[
            MessageHandler(filters.Regex("^➕ Navbat berildi$"), menu),
            MessageHandler(filters.Regex("^💰 Shoferdan olindi$"), menu),
        ],
        states={
            ASK_DATE: [
                CallbackQueryHandler(date_callback, pattern=r"^date:"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, custom_date),
            ],
            ASK_TRUCK: [
                CallbackQueryHandler(truck_callback, pattern=r"^truck:")
            ],
            ASK_AMOUNT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, amount_message)
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        per_chat=True,
        per_user=True,
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv)
    app.add_handler(MessageHandler(filters.Regex("^(📊 Hisobot|📥 Excel)$"), menu))

    app.run_webhook(
        listen="0.0.0.0",
        port=PORT,
        url_path=WEBHOOK_PATH,
        webhook_url=f"{WEBHOOK_URL}/{WEBHOOK_PATH}",
        drop_pending_updates=True,
    )

if __name__ == "__main__":
    main()
