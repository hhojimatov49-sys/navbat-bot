import os
import sqlite3
import threading
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from telegram import Update, ReplyKeyboardMarkup, InlineKeyboardMarkup, InlineKeyboardButton
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    ContextTypes, ConversationHandler, filters
)

TOKEN = os.environ["BOT_TOKEN"]
DB_PATH = os.environ.get("DB_PATH", "navbat.db")
PORT = int(os.environ.get("PORT", "10000"))

MENU = [
    ["➕ Navbat puli berildi", "💰 Shoferdan olindi"],
    ["📊 Hisobot"],
]

class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write(b"OK")
    def log_message(self, format, *args):
        return

def start_health_server():
    server = ThreadingHTTPServer(("0.0.0.0", PORT), HealthHandler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    print(f"Health server listening on 0.0.0.0:{PORT}")

def db():
    con = sqlite3.connect(DB_PATH)
    con.execute("""
        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            given_date TEXT NOT NULL,
            car TEXT NOT NULL,
            given_amount REAL NOT NULL,
            received_date TEXT,
            received_amount REAL
        )
    """)
    con.commit()
    return con

def cars():
    return ["60059KBA", "603181AA", "60227SBA", "604964AA"]

def fmt(n):
    return f"{n:,.2f}".rstrip("0").rstrip(".")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb = ReplyKeyboardMarkup(MENU, resize_keyboard=True)
    await update.message.reply_text(
        "Elektron navbat hisob-kitob botiga xush kelibsiz.",
        reply_markup=kb
    )

async def choose_given(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    context.user_data["mode"] = "given"
    await update.message.reply_text(
        "📅 Berilgan sanani kiriting (KK.OO.YYYY):\nMasalan: 29.08.2026"
    )
    return 1

async def choose_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    context.user_data["mode"] = "received"
    await update.message.reply_text(
        "📅 Olingan sanani kiriting (KK.OO.YYYY):\nMasalan: 30.08.2026"
    )
    return 1

async def date_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    try:
        dt = datetime.strptime(text, "%d.%m.%Y")
    except ValueError:
        await update.message.reply_text("Sana noto'g'ri. Masalan: 29.08.2026")
        return 1

    context.user_data["date"] = dt.strftime("%d.%m.%Y")
    mode = context.user_data["mode"]
    buttons = [
        [InlineKeyboardButton(c, callback_data=f"{mode}_car:{c}")]
        for c in cars()
    ]
    await update.message.reply_text(
        "🚛 Mashinani tanlang:",
        reply_markup=InlineKeyboardMarkup(buttons)
    )
    return 2

async def car_choice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    mode, car = q.data.split(":", 1)
    context.user_data["car"] = car
    context.user_data["mode"] = "given" if mode.startswith("given") else "received"
    await q.edit_message_text(
        f"🚛 Mashina: {car}\n💰 Summani kiriting:"
    )
    return 3

async def amount_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        amount = float(update.message.text.replace(" ", "").replace(",", "."))
        if amount < 0:
            raise ValueError
    except ValueError:
        await update.message.reply_text("Summani raqamda kiriting.")
        return 3

    mode = context.user_data["mode"]
    date = context.user_data["date"]
    car = context.user_data["car"]

    con = db()
    if mode == "given":
        con.execute(
            "INSERT INTO payments(given_date, car, given_amount) VALUES(?,?,?)",
            (date, car, amount)
        )
        msg = f"✅ Saqlandi:\nSana: {date}\nMashina: {car}\nBerilgan: {fmt(amount)}"
    else:
        row = con.execute(
            """SELECT id, given_amount FROM payments
               WHERE car=? AND received_amount IS NULL
               ORDER BY id DESC LIMIT 1""", (car,)
        ).fetchone()

        if not row:
            con.close()
            await update.message.reply_text(
                f"⚠️ {car} uchun hali 'berilgan' navbat puli topilmadi."
            )
            return ConversationHandler.END

        con.execute(
            "UPDATE payments SET received_date=?, received_amount=? WHERE id=?",
            (date, amount, row[0])
        )
        diff = amount - row[1]
        msg = (
            f"✅ Saqlandi:\nSana: {date}\nMashina: {car}\n"
            f"Olingan: {fmt(amount)}\nFarqi: {fmt(diff)}"
        )

    con.commit()
    con.close()
    await update.message.reply_text(msg)
    return ConversationHandler.END

async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    con = db()
    row = con.execute("""
        SELECT COALESCE(SUM(given_amount),0),
               COALESCE(SUM(received_amount),0),
               COALESCE(SUM(CASE WHEN received_amount IS NOT NULL
                                 THEN received_amount-given_amount ELSE 0 END),0)
        FROM payments
    """).fetchone()
    pending = con.execute(
        "SELECT COUNT(*) FROM payments WHERE received_amount IS NULL"
    ).fetchone()[0]
    con.close()

    await update.message.reply_text(
        "📊 Umumiy hisobot\n\n"
        f"Berilgan jami: {fmt(row[0])}\n"
        f"Olingan jami: {fmt(row[1])}\n"
        f"Farqi: {fmt(row[2])}\n"
        f"Olinmagan yozuvlar: {pending}"
    )

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("Bekor qilindi.")
    return ConversationHandler.END

def main():
    db().close()
    start_health_server()

    app = Application.builder().token(TOKEN).build()

    given_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^➕ Navbat puli berildi$"), choose_given)],
        states={
            1: [MessageHandler(filters.TEXT & ~filters.COMMAND, date_input)],
            2: [CallbackQueryHandler(car_choice, pattern=r"^given_car:")],
            3: [MessageHandler(filters.TEXT & ~filters.COMMAND, amount_input)],
        },
        fallbacks=[CommandHandler("cancel", cancel)]
    )

    received_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^💰 Shoferdan olindi$"), choose_received)],
        states={
            1: [MessageHandler(filters.TEXT & ~filters.COMMAND, date_input)],
            2: [CallbackQueryHandler(car_choice, pattern=r"^received_car:")],
            3: [MessageHandler(filters.TEXT & ~filters.COMMAND, amount_input)],
        },
        fallbacks=[CommandHandler("cancel", cancel)]
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(given_conv)
    app.add_handler(received_conv)
    app.add_handler(MessageHandler(filters.Regex("^📊 Hisobot$"), report))

    print("Bot started")
    app.run_polling()

if __name__ == "__main__":
    main()
