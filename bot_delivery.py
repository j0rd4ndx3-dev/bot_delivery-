# bot_delivery.py  — Versión: cierre con /total y reinicio automático al nuevo monto
import logging
import os
import re
from datetime import datetime
import pandas as pd
from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, CommandHandler, MessageHandler, filters

# ---------------- CONFIG ----------------
TOKEN = os.environ.get("TOKEN")
if not TOKEN:
    raise SystemExit("ERROR: la variable de entorno TOKEN no está definida. Define TOKEN antes de ejecutar el bot.")

EXCEL_FILE = "ganancias_delivery.xlsx"
# ----------------------------------------

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

ENTRIES_SHEET = "entries"
SESSIONS_SHEET = "sessions"

ENTRIES_COLS = ["SessionID", "UserID", "FechaHora", "EntregaID_Session", "Monto", "Nota"]
SESSIONS_COLS = ["SessionID", "UserID", "StartTime", "EndTime", "Active"]

# ---------- Helpers para Excel / almacenamiento ----------
def ensure_excel_exists():
    if not os.path.exists(EXCEL_FILE):
        df_e = pd.DataFrame(columns=ENTRIES_COLS)
        df_s = pd.DataFrame(columns=SESSIONS_COLS)
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
            df_e.to_excel(writer, sheet_name=ENTRIES_SHEET, index=False)
            df_s.to_excel(writer, sheet_name=SESSIONS_SHEET, index=False)

def read_entries():
    ensure_excel_exists()
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=ENTRIES_SHEET)
    except Exception:
        df = pd.DataFrame(columns=ENTRIES_COLS)
    return df

def read_sessions():
    ensure_excel_exists()
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SESSIONS_SHEET)
    except Exception:
        df = pd.DataFrame(columns=SESSIONS_COLS)
    return df

def write_all(entries_df, sessions_df):
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
        entries_df.to_excel(writer, sheet_name=ENTRIES_SHEET, index=False)
        sessions_df.to_excel(writer, sheet_name=SESSIONS_SHEET, index=False)

# ---------- Sesiones y operaciones ----------
def get_active_session_id(user_id: int):
    sessions = read_sessions()
    active = sessions[(sessions["UserID"] == user_id) & (sessions["Active"] == True)]
    if active.empty:
        return None
    return active.iloc[-1]["SessionID"]

def start_session_for_user(user_id: int) -> str:
    sessions = read_sessions()
    # If there is already an active session, return it
    active = sessions[(sessions["UserID"] == user_id) & (sessions["Active"] == True)]
    if not active.empty:
        return active.iloc[-1]["SessionID"]
    now = datetime.now().strftime("%Y%m%d%H%M%S")
    sid = f"{user_id}_{now}"
    new_row = {
        "SessionID": sid,
        "UserID": user_id,
        "StartTime": datetime.now().isoformat(sep=" ", timespec="seconds"),
        "EndTime": "",
        "Active": True
    }
    sessions = pd.concat([sessions, pd.DataFrame([new_row])], ignore_index=True)
    write_all(read_entries(), sessions)
    return sid

def end_session_for_user(user_id: int) -> (str, float, int):
    sessions = read_sessions()
    active = sessions[(sessions["UserID"] == user_id) & (sessions["Active"] == True)]
    if active.empty:
        return None, 0.0, 0
    sid = active.iloc[-1]["SessionID"]
    entries = read_entries()
    sub = entries[entries["SessionID"] == sid]
    total = 0.0
    count = 0
    if not sub.empty:
        total = float(sub["Monto"].sum())
        count = sub.shape[0]
    idx = sessions[(sessions["SessionID"] == sid)].index
    sessions.loc[idx, "EndTime"] = datetime.now().isoformat(sep=" ", timespec="seconds")
    sessions.loc[idx, "Active"] = False
    write_all(entries, sessions)
    return sid, total, count

def append_entry_for_session(session_id: str, user_id: int, monto: float, nota: str):
    entries = read_entries()
    sub = entries[entries["SessionID"] == session_id]
    next_id = 1
    if not sub.empty:
        next_id = int(sub["EntregaID_Session"].max()) + 1
    new = {
        "SessionID": session_id,
        "UserID": user_id,
        "FechaHora": datetime.now().isoformat(sep=" ", timespec="seconds"),
        "EntregaID_Session": next_id,
        "Monto": round(float(monto), 2),
        "Nota": nota
    }
    entries = pd.concat([entries, pd.DataFrame([new])], ignore_index=True)
    write_all(entries, read_sessions())
    return next_id

def list_entries_for_session(session_id: str):
    entries = read_entries()
    sub = entries[entries["SessionID"] == session_id].sort_values("EntregaID_Session")
    return sub

def edit_entry(session_id: str, entrega_n: int, nuevo_monto: float, nota: str = None):
    entries = read_entries()
    mask = (entries["SessionID"] == session_id) & (entries["EntregaID_Session"] == entrega_n)
    if not mask.any():
        return False
    idx = entries[mask].index
    entries.loc[idx, "Monto"] = round(float(nuevo_monto), 2)
    if nota is not None:
        entries.loc[idx, "Nota"] = nota
    write_all(entries, read_sessions())
    return True

def delete_entry(session_id: str, entrega_n: int):
    entries = read_entries()
    mask = (entries["SessionID"] == session_id) & (entries["EntregaID_Session"] == entrega_n)
    if not mask.any():
        return False
    entries = entries[~mask].copy()
    sub_mask = entries["SessionID"] == session_id
    if sub_mask.any():
        sub = entries[sub_mask].sort_values("FechaHora").reset_index(drop=True)
        sub["EntregaID_Session"] = range(1, len(sub) + 1)
        entries = pd.concat([entries[~sub_mask], sub], ignore_index=True)
    write_all(entries, read_sessions())
    return True

# ---------- Utilidades ----------
def parse_amount(text: str):
    m = re.search(r'[-+]?\d+[.,]?\d*', text)
    if not m:
        return None
    s = m.group(0).replace(",", ".")
    try:
        return float(s)
    except:
        return None

# ---------- Handlers ----------
async def start_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "Bot listo ✅\n\nComandos:\n"
        "/startday - Inicia jornada\n"
        "/total or /endday - Cierra jornada activa y muestra total\n"
        "/list - Lista entregas de la sesión activa\n"
        "/editar N monto - Edita la entrega N de la sesión activa\n"
        "/borrar N - Borra la entrega N de la sesión activa\n"
        "/excel - Te envío el archivo con todo el historial\n"
        "/help - Mostrar ayuda\n\n"
        "También puedes simplemente enviar el monto (ej: 8.50). Si no hay sesión activa, el bot iniciará una nueva automáticamente."
    )
    await update.message.reply_text(text)

async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await start_cmd(update, context)

async def startday_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    sid = start_session_for_user(user.id)
    await update.message.reply_text(f"✅ Jornada iniciada. SessionID: {sid}\nEnvía montos y se guardarán en esta sesión.")

async def list_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    sid = get_active_session_id(user.id)
    if not sid:
        await update.message.reply_text("No tienes una sesión activa.")
        return
    sub = list_entries_for_session(sid)
    if sub.empty:
        await update.message.reply_text("Aún no hay entregas en la sesión activa.")
        return
    lines = []
    for _, row in sub.iterrows():
        lines.append(f"#{int(row['EntregaID_Session'])} — S/ {float(row['Monto']):.2f}  ({row['FechaHora']})")
    text = "Entregas sesión activa:\n" + "\n".join(lines)
    await update.message.reply_text(text)

async def total_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    sid = get_active_session_id(user.id)
    if not sid:
        await update.message.reply_text("No hay sesión activa que cerrar.")
        return
    sid_closed, total, count = end_session_for_user(user.id)
    await update.message.reply_text(f"✅ Jornada finalizada.\nSession: {sid_closed}\nEntregas: {count}\nGanancia total: S/ {total:.2f}")

async def send_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ensure_excel_exists()
    await update.message.reply_document(open(EXCEL_FILE, "rb"))

async def editar_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    sid = get_active_session_id(user.id)
    if not sid:
        await update.message.reply_text("No tienes una sesión activa.")
        return
    text = update.message.text.strip()
    parts = text.split()
    if len(parts) < 3:
        await update.message.reply_text("Uso: /editar N nuevo_monto (ej: /editar 3 8.50)")
        return
    try:
        n = int(parts[1])
        nuevo = parse_amount(" ".join(parts[2:]))
        if nuevo is None:
            await update.message.reply_text("No detecté un monto válido.")
            return
    except ValueError:
        await update.message.reply_text("El primer parámetro debe ser el número de entrega (ej: 1).")
        return
    nota = " ".join(parts[3:]) if len(parts) > 3 else None
    ok = edit_entry(sid, n, nuevo, nota)
    if not ok:
        await update.message.reply_text(f"No encontré la entrega #{n} en la sesión activa.")
        return
    await update.message.reply_text(f"✔ Entrega #{n} actualizada a S/ {nuevo:.2f}")

async def borrar_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    sid = get_active_session_id(user.id)
    if not sid:
        await update.message.reply_text("No tienes una sesión activa.")
        return
    text = update.message.text.strip()
    parts = text.split()
    if len(parts) != 2:
        await update.message.reply_text("Uso: /borrar N  (ej: /borrar 4)")
        return
    try:
        n = int(parts[1])
    except ValueError:
        await update.message.reply_text("El parámetro N debe ser un número (ej: /borrar 2).")
        return
    ok = delete_entry(sid, n)
    if not ok:
        await update.message.reply_text(f"No encontré la entrega #{n} en la sesión activa.")
        return
    await update.message.reply_text(f"🗑 Entrega #{n} eliminada. Las demás entregas se renumeraron.")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text.strip()
    user = update.effective_user
    if txt.startswith("/"):
        await update.message.reply_text("Comando desconocido. Usa /help para ver comandos.")
        return
    amount = parse_amount(txt)
    if amount is None:
        await update.message.reply_text("No detecté un monto válido. Envía algo como: 8.5 o 'entrega 8,50'.")
        return
    # Si no hay sesión activa: iniciarla automáticamente (comportamiento que pediste)
    sid = get_active_session_id(user.id)
    if not sid:
        sid = start_session_for_user(user.id)
        # informar al usuario que inició sesión automáticamente
        await update.message.reply_text("He creado automáticamente una nueva jornada para ti (se inició al registrar este monto).")
    entrega_id = append_entry_for_session(sid, user.id, amount, txt)
    sub = list_entries_for_session(sid)
    total = float(sub["Monto"].sum()) if not sub.empty else 0.0
    await update.message.reply_text(f"Registrado: entrega #{entrega_id} — S/ {amount:.2f}\nTotal sesión: S/ {total:.2f}")

# ---------- Main ----------
def main():
    ensure_excel_exists()
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start_cmd))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("startday", startday_cmd))
    app.add_handler(CommandHandler("list", list_cmd))
    app.add_handler(CommandHandler("total", total_cmd))
    app.add_handler(CommandHandler("endday", total_cmd))
    app.add_handler(CommandHandler("excel", send_excel))
    app.add_handler(CommandHandler("editar", editar_cmd))
    app.add_handler(CommandHandler("borrar", borrar_cmd))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))

    print("Bot corriendo... Ctrl+C para detener")
    app.run_polling()

if __name__ == "__main__":
    main()
