# bot_delivery.py
import logging
from datetime import datetime, date
import re
import os
import pandas as pd
from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, CommandHandler, MessageHandler, filters

# ---------- CONFIG ----------
TOKEN = "7988204764:AAFX5kMW--DXO9IBYClfF-PcG0PXoS8bGoA" # <- pega aquí el token que te dio @BotFather
EXCEL_FILE = "ganancias_delivery.xlsx"
# ----------------------------

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

COLUMNS = ["Fecha", "Hora", "EntregaID_Dia", "Monto", "Nota"]

def ensure_excel_exists():
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(EXCEL_FILE, index=False)

def read_excel():
    ensure_excel_exists()
    return pd.read_excel(EXCEL_FILE)

def append_entry(entry: dict):
    df = read_excel()
    df = pd.concat([df, pd.DataFrame([entry])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

def parse_amount(text: str):
    # busca un número con punto o coma, opcional signo +
    m = re.search(r'[-+]?\d+[.,]?\d*', text)
    if not m:
        return None
    s = m.group(0).replace(',', '.')
    try:
        return float(s)
    except:
        return None

def get_today_count():
    df = read_excel()
    today_str = date.today().strftime("%Y-%m-%d")
    if df.empty:
        return 0
    return df[df["Fecha"] == today_str].shape[0]

def get_today_total():
    df = read_excel()
    today_str = date.today().strftime("%Y-%m-%d")
    if df.empty:
        return 0.0
    sub = df[df["Fecha"] == today_str]
    if sub.empty:
        return 0.0
    return float(sub["Monto"].sum())

# ---------- Handlers ----------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "¡Listo! Envía el monto de cada entrega (ej: 8.5 o +8.50 o entrega 8,5).\n\n"
        "Comandos:\n"
        "/total - Muestra la suma del día\n"
        "/excel - Te envío el archivo con todo el historial\n"
        "/help - Ayuda\n\n"
        "El archivo se guarda en el servidor como: " + EXCEL_FILE
    )
    await update.message.reply_text(text)

async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await start(update, context)

async def send_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ensure_excel_exists()
    await update.message.reply_document(open(EXCEL_FILE, "rb"))

async def total_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    total = get_today_total()
    count = get_today_count()
    await update.message.reply_text(f"Entregas hoy: {count}\nGanancia total hoy: {total:.2f}")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text.strip()
    amount = parse_amount(txt)
    if amount is None:
        await update.message.reply_text("No detecté un monto válido. Envía algo como: 8.5 o entrega 6.20")
        return

    # construir entrada
    now = datetime.now()
    fecha = now.strftime("%Y-%m-%d")
    hora = now.strftime("%H:%M:%S")
    entrega_id = get_today_count() + 1
    nota = txt

    entry = {
        "Fecha": fecha,
        "Hora": hora,
        "EntregaID_Dia": entrega_id,
        "Monto": round(amount, 2),
        "Nota": nota
    }
    try:
        append_entry(entry)
    except Exception as e:
        logger.exception("Error guardando entrada")
        await update.message.reply_text("Ocurrió un error guardando el registro.")
        return

    total = get_today_total()
    await update.message.reply_text(f"Registrado: entrega #{entrega_id} — {amount:.2f}\nTotal hoy: {total:.2f}")

# ---------- Main ----------
def main():
    ensure_excel_exists()
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("excel", send_excel))
    app.add_handler(CommandHandler("total", total_cmd))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))

    print("Bot corriendo... Ctrl+C para detener")
    app.run_polling()

if __name__ == "__main__":
    main()
