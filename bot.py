import os
import logging
from dotenv import load_dotenv

from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    filters,
    ContextTypes
)

from database import (
    init_db,
    insert_orders,
    search,
    produk_summary,
    toko_summary,
    status
)

from parser import parse_line
from exporter import export_excel
from dashboard import dashboard


load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")

logging.basicConfig(level=logging.INFO)


# ================= START =================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):

    await update.message.reply_text(
        "📊 POS ENTERPRISE BOT",
        reply_markup=dashboard()
    )


# ================= FILE UPLOAD =================

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):

    file = await update.message.document.get_file()

    path = "upload.txt"

    await file.download_to_drive(path)

    rows = []

    with open(path,"r",encoding="utf8",errors="ignore") as f:

        for line in f:

            r = parse_line(line)

            if r:
                rows.extend(r)

    if rows:

        insert_orders(rows)

    await update.message.reply_text(
        f"✅ Upload selesai\nData masuk : {len(rows)}"
    )


# ================= SEARCH =================

async def cmd_search(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if not context.args:

        await update.message.reply_text("Gunakan /search nama_atau_oid")
        return

    rows = search(context.args[0])

    text="🔎 HASIL SEARCH\n\n"

    for r in rows:

        text+=f"""
Nama : {r[0]}
OID : {r[1]}
Produk : {r[2]}
Qty : {r[3]}
Toko : {r[4]}

"""

    await update.message.reply_text(text)


# ================= CALLBACK =================

async def menu(update: Update, context: ContextTypes.DEFAULT_TYPE):

    q = update.callback_query
    await q.answer()

    if q.data=="menu|produk":

        rows=produk_summary()

        text="📦 REKAP PRODUK\n\n"

        for i,r in enumerate(rows,1):

            text+=f"{i}. {r[0]} — {r[1]} pcs\n"

        await q.edit_message_text(text)

    if q.data=="menu|toko":

        rows=toko_summary()

        text="🏪 REKAP TOKO\n\n"

        for i,r in enumerate(rows,1):

            text+=f"{i}. {r[0]} — {r[1]} pcs\n"

        await q.edit_message_text(text)

    if q.data=="menu|excel":

        file=export_excel()

        await context.bot.send_document(
            chat_id=q.message.chat_id,
            document=file,
            filename="report.xlsx"
        )

    if q.data=="menu|status":

        total=status()

        await q.edit_message_text(
            f"📡 STATUS SERVER\nOrders : {total}"
        )

    if q.data=="menu|help":

        await q.edit_message_text(
            """
❓ BANTUAN

Upload TXT → kirim file txt
/search nama
/search OID

Dashboard tersedia
"""
        )


# ================= MAIN =================

def main():

    init_db()

    app=Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start",start))
    app.add_handler(CommandHandler("search",cmd_search))

    app.add_handler(
        MessageHandler(filters.Document.TEXT,handle_file)
    )

    app.add_handler(
        CallbackQueryHandler(menu,pattern="menu")
    )

    print("BOT RUNNING")

    app.run_polling()


if __name__=="__main__":
    main()