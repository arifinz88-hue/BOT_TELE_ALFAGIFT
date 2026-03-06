import re
import sqlite3
import hashlib
import pandas as pd
import os
import threading

from io import BytesIO
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QPushButton, QLabel, QDateEdit, QProgressBar, QTableWidget,
    QTableWidgetItem, QFileDialog, QMessageBox
)
from PySide6.QtCore import QThread, Signal, Qt, QDate, QTimer

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, ContextTypes

DB = "rekap_cache.db"

PRODUCT_RE = re.compile(r"Produk=\s*(.*?)\s*Qty=\s*(\d+)", re.I)
OID_RE = re.compile(r"O-(\d{6})-([A-Z0-9]+)", re.I)
VALID_STATUS = ("SELESAI", "SEDANG DIPROSES")

tg_action_map = {}
chat_range = {}


# ================= DB =================
def init_db():
    conn = sqlite3.connect(DB)
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS orders(
            oid TEXT,
            tanggal TEXT,
            toko TEXT,
            nama TEXT,
            produk TEXT,
            qty INTEGER,
            UNIQUE(oid, produk)
        )
    """)

    try:
        cols = [r[1] for r in c.execute("PRAGMA table_info(orders)").fetchall()]
        if "nama" not in cols:
            c.execute("ALTER TABLE orders ADD COLUMN nama TEXT")
    except:
        pass

    c.execute("""
        CREATE TABLE IF NOT EXISTS scanned_files(
            path TEXT PRIMARY KEY,
            hash TEXT
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS product_status(
            toko TEXT,
            tanggal TEXT,
            produk TEXT,
            is_taken INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            taken_at TEXT,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (toko, tanggal, produk)
        )
    """)

    try:
        cols = [r[1] for r in c.execute("PRAGMA table_info(product_status)").fetchall()]
        if "created_at" not in cols:
            c.execute("ALTER TABLE product_status ADD COLUMN created_at TEXT")
        if "taken_at" not in cols:
            c.execute("ALTER TABLE product_status ADD COLUMN taken_at TEXT")
        if "updated_at" not in cols:
            c.execute("ALTER TABLE product_status ADD COLUMN updated_at TEXT")
    except:
        pass

    conn.commit()
    conn.close()


def file_hash(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def turbo_scan(folder, progress=None):
    init_db()
    conn = sqlite3.connect(DB)
    c = conn.cursor()

    files = list(Path(folder).rglob("*.txt"))
    total = len(files)
    inserted = 0

    for i, path in enumerate(files):
        if progress:
            progress(i + 1, total)

        path = str(path)
        h = file_hash(path)

        old = c.execute("SELECT hash FROM scanned_files WHERE path=?", (path,)).fetchone()
        if old and old[0] == h:
            continue

        batch = []

        with open(path, "r", encoding="utf8", errors="ignore") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue

                up = line.upper()
                if not any(s in up for s in VALID_STATUS):
                    continue

                nama = line.split(":", 1)[0].strip() if ":" in line else "-"
                if not nama:
                    nama = "-"

                m = OID_RE.search(line)
                if not m:
                    continue

                tanggal = m.group(1)
                oid = f"O-{tanggal}-{m.group(2)}"

                toko = "UNKNOWN"
                if "|:" in line:
                    meta = line.split("|:", 1)[1].split(":")
                    if len(meta) >= 3:
                        toko = f"{meta[1].strip()} - {meta[2].strip()}"

                for p in PRODUCT_RE.finditer(line):
                    produk = p.group(1).strip()
                    qty = int(p.group(2))
                    batch.append((oid, tanggal, toko, nama, produk, qty))

        if batch:
            c.executemany(
                "INSERT OR IGNORE INTO orders (oid,tanggal,toko,nama,produk,qty) VALUES (?,?,?,?,?,?)",
                batch
            )
            inserted += len(batch)

            for _, tanggal, toko, _, produk, _ in batch:
                c.execute("""
                    INSERT OR IGNORE INTO product_status
                    (toko, tanggal, produk, is_taken, created_at, updated_at)
                    VALUES (?,?,?,0,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)
                """, (toko, tanggal, produk))

        c.execute("INSERT OR REPLACE INTO scanned_files VALUES (?,?)", (path, h))
        conn.commit()

    conn.close()
    return inserted


def set_taken_status(toko, tanggal, produk, status):
    conn = sqlite3.connect(DB)

    if int(status) == 1:
        conn.execute("""
            INSERT INTO product_status
            (toko, tanggal, produk, is_taken, created_at, taken_at, updated_at)
            VALUES (?,?,?,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)
            ON CONFLICT(toko, tanggal, produk)
            DO UPDATE SET
                is_taken=1,
                taken_at=CURRENT_TIMESTAMP,
                updated_at=CURRENT_TIMESTAMP
        """, (toko, tanggal, produk))
    else:
        conn.execute("""
            INSERT INTO product_status
            (toko, tanggal, produk, is_taken, created_at, taken_at, updated_at)
            VALUES (?,?,?,0,CURRENT_TIMESTAMP,NULL,CURRENT_TIMESTAMP)
            ON CONFLICT(toko, tanggal, produk)
            DO UPDATE SET
                is_taken=0,
                taken_at=NULL,
                updated_at=CURRENT_TIMESTAMP
        """, (toko, tanggal, produk))

    conn.commit()
    conn.close()


# ================= THREAD LOADER =================
class Loader(QThread):
    progress = Signal(int, int)
    finished = Signal(int)

    def __init__(self, folder):
        super().__init__()
        self.folder = folder

    def run(self):
        n = turbo_scan(self.folder, self.progress.emit)
        self.finished.emit(n)


# ================= REPORT HELPERS =================
def build_report_rows_for_store(toko, t1, t2):
    conn = sqlite3.connect(DB)
    rows = conn.execute("""
        SELECT o.tanggal,
               o.produk,
               SUM(o.qty) as total_qty,
               COALESCE(ps.is_taken, 0) as is_taken,
               COALESCE(ps.created_at, '-') as created_at,
               COALESCE(ps.taken_at, '-') as taken_at
        FROM orders o
        LEFT JOIN product_status ps
          ON ps.toko = o.toko AND ps.tanggal = o.tanggal AND ps.produk = o.produk
        WHERE o.toko=? AND o.tanggal BETWEEN ? AND ?
        GROUP BY o.tanggal, o.produk, ps.is_taken, ps.created_at, ps.taken_at
        ORDER BY o.tanggal, total_qty DESC, o.produk
    """, (toko, t1, t2)).fetchall()
    conn.close()
    return rows


def build_wa_message_for_store(toko, t1, t2):
    rows = build_report_rows_for_store(toko, t1, t2)
    if not rows:
        return None

    by_tgl = {}
    for tanggal, produk, qty, is_taken, created_at, taken_at in rows:
        by_tgl.setdefault(tanggal, []).append(
            (produk, int(qty), int(is_taken), created_at, taken_at)
        )

    msg = []
    msg.append("==========================================")
    msg.append(f"REPORT TOKO : {toko}")
    msg.append(f"RANGE       : {t1} s/d {t2}")
    msg.append("")

    grand_total = 0

    for tgl in sorted(by_tgl):
        msg.append(f"TANGGAL ORDER : {tgl}")
        for i, (produk, qty, is_taken, created_at, taken_at) in enumerate(by_tgl[tgl], 1):
            tanda = "✅" if is_taken else "⬜"
            t_ambil = taken_at if taken_at and taken_at != "-" else "BELUM DIAMBIL"

            msg.append(f"{i}. {tanda} {produk}")
            msg.append(f"   QTY       : {qty} pcs")
            msg.append(f"   TGL BUAT  : {created_at}")
            msg.append(f"   TGL AMBIL : {t_ambil}")
            grand_total += qty
        msg.append("")

    msg.append("==========================================")
    msg.append(f"TOTAL QTY : {grand_total} pcs")
    msg.append("Laporan otomatis ASP_GROUP")
    msg.append("==========================================")
    return "\n".join(msg)


# ================= TELEGRAM DB =================
def tg_db_list_toko(t1, t2):
    conn = sqlite3.connect(DB)
    rows = conn.execute("""
        SELECT toko, COUNT(DISTINCT oid) as trx
        FROM orders
        WHERE tanggal BETWEEN ? AND ?
        GROUP BY toko
        ORDER BY trx DESC
    """, (t1, t2)).fetchall()
    conn.close()
    return rows


def tg_db_rekap_produk(toko, t1, t2):
    conn = sqlite3.connect(DB)
    rows = conn.execute("""
        SELECT o.tanggal,
               o.produk,
               SUM(o.qty) as total_qty,
               COALESCE(ps.is_taken, 0) as is_taken,
               COALESCE(ps.created_at, '-') as created_at,
               COALESCE(ps.taken_at, '-') as taken_at
        FROM orders o
        LEFT JOIN product_status ps
          ON ps.toko = o.toko AND ps.tanggal = o.tanggal AND ps.produk = o.produk
        WHERE o.toko=? AND o.tanggal BETWEEN ? AND ?
        GROUP BY o.tanggal, o.produk, ps.is_taken, ps.created_at, ps.taken_at
        ORDER BY o.tanggal, total_qty DESC, o.produk
    """, (toko, t1, t2)).fetchall()
    conn.close()
    return rows


def tg_db_nama_oid(toko, t1, t2):
    conn = sqlite3.connect(DB)
    rows = conn.execute("""
        SELECT tanggal, COALESCE(NULLIF(nama,''), '-') as nama, oid
        FROM orders
        WHERE toko=? AND tanggal BETWEEN ? AND ?
        GROUP BY tanggal, nama, oid
        ORDER BY tanggal, oid
    """, (toko, t1, t2)).fetchall()
    conn.close()
    return rows


def tg_chunk(s, n=3500):
    return [s[i:i + n] for i in range(0, len(s), n)]


def tg_make_token(prefix, *parts):
    raw = "|".join(str(x) for x in parts)
    token = hashlib.md5(raw.encode("utf-8")).hexdigest()[:16]
    tg_action_map[token] = (prefix, parts)
    return token


# ================= TELEGRAM UI HELPERS =================
async def tg_show_store_list(chat_id, t1, t2, context, edit_message=None):
    rows = tg_db_list_toko(t1, t2)
    if not rows:
        text = "❌ Tidak ada data pada range itu."
        if edit_message:
            await edit_message.edit_message_text(text)
        else:
            await context.bot.send_message(chat_id=chat_id, text=text)
        return

    buttons = []
    lines = []
    lines.append("🏪 <b>DAFTAR TOKO</b>")
    lines.append(f"<i>Range: {t1} s/d {t2}</i>")
    lines.append("")

    for i, (toko, trx) in enumerate(rows, 1):
        lines.append(f"{i}. <b>{toko}</b> | Trx: {trx}")
        buttons.append([
            InlineKeyboardButton(
                f"{toko} | Trx:{trx}",
                callback_data=f"toko|{toko}"
            )
        ])

    text = "\n".join(lines)

    if edit_message:
        await edit_message.edit_message_text(
            text,
            reply_markup=InlineKeyboardMarkup(buttons),
            parse_mode="HTML"
        )
    else:
        await context.bot.send_message(
            chat_id=chat_id,
            text=text,
            reply_markup=InlineKeyboardMarkup(buttons),
            parse_mode="HTML"
        )


async def send_report_to_telegram(chat_id, toko, t1, t2, context):
    rows = build_report_rows_for_store(toko, t1, t2)
    if not rows:
        await context.bot.send_message(chat_id=chat_id, text="❌ Tidak ada report.")
        return

    by_tgl = {}
    for tanggal, produk, qty, is_taken, created_at, taken_at in rows:
        by_tgl.setdefault(tanggal, []).append(
            (produk, int(qty), int(is_taken), created_at, taken_at)
        )

    out = []
    out.append("📦 <b>REPORT TOKO</b>")
    out.append(f"<b>{toko}</b>")
    out.append(f"<i>Range: {t1} s/d {t2}</i>")
    out.append("")

    excel_rows = []
    grand_total = 0

    for tgl in sorted(by_tgl):
        out.append(f"<i>Tanggal Order: {tgl}</i>")
        for i, (produk, qty, is_taken, created_at, taken_at) in enumerate(by_tgl[tgl], 1):
            status_txt = "🟢 <b>SUDAH DIAMBIL</b>" if is_taken else "🔴 <b>BELUM DIAMBIL</b>"
            t_ambil = taken_at if taken_at and taken_at != "-" else "-"

            out.append(f"{i}. <b>{produk}</b>")
            out.append(f"   Qty          : {qty} pcs")
            out.append(f"   Tgl Buat     : {created_at}")
            out.append(f"   Tgl Ambil    : {t_ambil}")
            out.append(f"   Status       : {status_txt}")
            out.append("")

            excel_rows.append({
                "Tanggal Order": tgl,
                "Toko": toko,
                "Produk": produk,
                "Total Qty": qty,
                "Status": "SUDAH DIAMBIL" if is_taken else "BELUM DIAMBIL",
                "Tanggal Buat": created_at,
                "Tanggal Ambil": t_ambil
            })
            grand_total += qty

    out.append(f"<b>TOTAL QTY : {grand_total} pcs</b>")
    text_report = "\n".join(out).strip()

    for part in tg_chunk(text_report):
        await context.bot.send_message(chat_id=chat_id, text=part, parse_mode="HTML")

    txt_bytes = BytesIO(text_report.encode("utf-8"))
    txt_bytes.name = f"REPORT_{toko}_{t1}_{t2}.txt".replace(" ", "_").replace("/", "-")
    await context.bot.send_document(chat_id=chat_id, document=txt_bytes)

    df = pd.DataFrame(excel_rows)
    xlsx_bytes = BytesIO()
    with pd.ExcelWriter(xlsx_bytes, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Report")
    xlsx_bytes.seek(0)
    xlsx_bytes.name = f"REPORT_{toko}_{t1}_{t2}.xlsx".replace(" ", "_").replace("/", "-")
    await context.bot.send_document(chat_id=chat_id, document=xlsx_bytes)


async def tg_render_toko_view(chat_id, query, toko, t1, t2, context):
    rows = tg_db_rekap_produk(toko, t1, t2)
    if not rows:
        await query.edit_message_text("❌ Data kosong.")
        return

    by_tgl = {}
    for tanggal, produk, qty, is_taken, created_at, taken_at in rows:
        by_tgl.setdefault(tanggal, []).append(
            (produk, int(qty), int(is_taken), created_at, taken_at)
        )

    out = []
    out.append("📦 <b>REPORT TOKO</b>")
    out.append(f"<b>{toko}</b>")
    out.append(f"<i>Range: {t1} s/d {t2}</i>")
    out.append("")

    buttons = []

    for tgl in sorted(by_tgl):
        out.append(f"<i>Tanggal Order: {tgl}</i>")
        total = 0

        for i, (produk, qty, is_taken, created_at, taken_at) in enumerate(by_tgl[tgl], 1):
            total += qty
            status_txt = "🟢 <b>SUDAH DIAMBIL</b>" if is_taken else "🔴 <b>BELUM DIAMBIL</b>"
            t_ambil = taken_at if taken_at and taken_at != "-" else "-"

            out.append(f"{i}. <b>{produk}</b>")
            out.append(f"   Qty          : {qty} pcs")
            out.append(f"   Tgl Buat     : {created_at}")
            out.append(f"   Tgl Ambil    : {t_ambil}")
            out.append(f"   Status       : {status_txt}")
            out.append("")

            tok_sudah = tg_make_token("toggle", toko, tgl, produk, 1)
            tok_belum = tg_make_token("toggle", toko, tgl, produk, 0)

            buttons.append([
                InlineKeyboardButton("✅ Sudah", callback_data=f"act|{tok_sudah}"),
                InlineKeyboardButton("⬜ Belum", callback_data=f"act|{tok_belum}")
            ])

        out.append(f"<b>TOTAL QTY : {total} pcs</b>")
        out.append("")

    tok_nama = tg_make_token("namaoid", toko, t1, t2)
    tok_done = tg_make_token("done_report", toko, t1, t2)
    tok_back = tg_make_token("back_store_list", t1, t2)

    buttons.append([
        InlineKeyboardButton("👥 Nama + Order ID", callback_data=f"act|{tok_nama}")
    ])
    buttons.append([
        InlineKeyboardButton("✅ Selesai & Kirim Report", callback_data=f"act|{tok_done}")
    ])
    buttons.append([
        InlineKeyboardButton("⬅ Back", callback_data=f"act|{tok_back}")
    ])

    text = "\n".join(out).strip()
    await query.edit_message_text(
        text,
        reply_markup=InlineKeyboardMarkup(buttons),
        parse_mode="HTML"
    )


# ================= TELEGRAM BOT =================
async def tg_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "✅ MENU7 siap.\n\n"
        "Format:\n"
        "/menu7 yyMMdd yyMMdd\n\n"
        "Contoh:\n"
        "/menu7 251124 251124"
    )


async def tg_menu7(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if len(context.args) != 2:
        await update.message.reply_text(
            "❌ Format: /menu7 yyMMdd yyMMdd\nContoh: /menu7 251124 251124"
        )
        return

    t1, t2 = context.args[0].strip(), context.args[1].strip()
    chat_id = update.message.chat_id
    chat_range[chat_id] = (t1, t2)

    await tg_show_store_list(chat_id, t1, t2, context)


async def tg_toko_click(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()

    chat_id = q.message.chat_id
    if chat_id not in chat_range:
        await q.edit_message_text("❌ Range belum ada. Pakai /menu7 dulu.")
        return

    _, toko = q.data.split("|", 1)
    toko = toko.strip()
    t1, t2 = chat_range[chat_id]

    await tg_render_toko_view(chat_id, q, toko, t1, t2, context)


async def tg_action_click(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()

    chat_id = q.message.chat_id
    if chat_id not in chat_range:
        await q.edit_message_text("❌ Range belum ada. Pakai /menu7 dulu.")
        return

    try:
        _, token = q.data.split("|", 1)
        action = tg_action_map.get(token)
        if not action:
            await q.edit_message_text("❌ Aksi expired. Ulangi /menu7.")
            return

        prefix, parts = action

        if prefix == "toggle":
            toko, tanggal, produk, next_status = parts
            next_status = int(next_status)
            set_taken_status(toko, tanggal, produk, next_status)
            t1, t2 = chat_range[chat_id]
            await tg_render_toko_view(chat_id, q, toko, t1, t2, context)
            return

        if prefix == "namaoid":
            toko, t1, t2 = parts
            rows = tg_db_nama_oid(toko, t1, t2)
            if not rows:
                await q.message.reply_text("❌ Data kosong.")
                return

            by_tgl = {}
            for tanggal, nama, oid in rows:
                by_tgl.setdefault(tanggal, []).append((nama, oid))

            out = []
            out.append("👥 <b>NAMA + ORDER ID</b>")
            out.append(f"<b>{toko}</b>")
            out.append(f"<i>Range: {t1} s/d {t2}</i>")
            out.append("")

            for tgl in sorted(by_tgl):
                out.append(f"<i>Tanggal: {tgl}</i>")
                for i, (nama, oid) in enumerate(by_tgl[tgl], 1):
                    out.append(f"{i}. <b>{nama}</b> | <code>{oid}</code>")
                out.append("")

            text = "\n".join(out).strip()
            parts_text = tg_chunk(text)
            tok_back_toko = tg_make_token("back_toko_view", toko, t1, t2)
            back_markup = InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅ Back", callback_data=f"act|{tok_back_toko}")]
            ])

            for idx, part in enumerate(parts_text):
                if idx == len(parts_text) - 1:
                    await context.bot.send_message(
                        chat_id=chat_id,
                        text=part,
                        parse_mode="HTML",
                        reply_markup=back_markup
                    )
                else:
                    await context.bot.send_message(
                        chat_id=chat_id,
                        text=part,
                        parse_mode="HTML"
                    )
            return

        if prefix == "back_store_list":
            t1, t2 = parts
            await tg_show_store_list(chat_id, t1, t2, context, edit_message=q)
            return

        if prefix == "back_toko_view":
            toko, t1, t2 = parts
            await tg_render_toko_view(chat_id, q, toko, t1, t2, context)
            return

        if prefix == "done_report":
            toko, t1, t2 = parts
            await q.edit_message_text("✅ Selesai. Kirim report ke Telegram...")
            await send_report_to_telegram(chat_id, toko, t1, t2, context)
            return

    except Exception as e:
        await q.edit_message_text(f"❌ Error: {e}")


def run_telegram_bot():
    token = os.environ.get("BOT_TOKEN", "").strip().replace(" ", "")
    if not token:
        print("BOT_TOKEN belum diset. Bot Telegram tidak jalan.")
        print("Set dulu di PowerShell: $env:BOT_TOKEN='123:ABC...'")
        return

    app = Application.builder().token(token).build()
    app.add_handler(CommandHandler("start", tg_start))
    app.add_handler(CommandHandler("menu7", tg_menu7))
    app.add_handler(CallbackQueryHandler(tg_toko_click, pattern=r"^toko\|"))
    app.add_handler(CallbackQueryHandler(tg_action_click, pattern=r"^act\|"))

    print("Telegram bot aktif... (/menu7)")
    app.run_polling()


# ================= GUI APP =================
class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("POS ENTERPRISE FINAL 🚀")
        self.resize(1400, 800)

        self.selected_toko = None
        self.mode = None
        self.last_folder = None

        self.init_ui()
        self.dark_mode()

        self.auto_timer = QTimer(self)
        self.auto_timer.timeout.connect(self.auto_reload)
        self.auto_timer.start(5 * 60 * 1000)

    def init_ui(self):
        main = QWidget()
        self.setCentralWidget(main)
        root = QHBoxLayout(main)

        side = QVBoxLayout()
        root.addLayout(side, 1)

        def btn(txt, fn):
            b = QPushButton(txt)
            b.clicked.connect(fn)
            b.setMinimumHeight(44)
            side.addWidget(b)

        btn("⚡ Load Data", self.load_data)
        btn("📅 Rekap Tanggal", self.rekap_tanggal)
        btn("🏪 Rekap Per Toko", self.rekap_toko)
        btn("📦 Rekap Semua Barang", self.rekap_semua_barang)
        btn("🧾 Export Report Toko", self.export_excel_report_toko)
        btn("📤 Kirim WA", self.kirim_group)
        btn("⬅ Back", self.go_back)
        btn("💾 Export Excel", self.export_excel)

        side.addStretch()

        content = QVBoxLayout()
        root.addLayout(content, 4)

        self.title = QLabel("READY")
        self.title.setAlignment(Qt.AlignCenter)

        date = QHBoxLayout()
        self.tgl1 = QDateEdit()
        self.tgl2 = QDateEdit()

        for t in (self.tgl1, self.tgl2):
            t.setCalendarPopup(True)
            t.setDisplayFormat("yyMMdd")
            t.setDate(QDate.currentDate())

        date.addWidget(QLabel("Dari"))
        date.addWidget(self.tgl1)
        date.addWidget(QLabel("Sampai"))
        date.addWidget(self.tgl2)

        self.progress = QProgressBar()
        self.table = QTableWidget()
        self.table.cellDoubleClicked.connect(self.table_click)

        content.addWidget(self.title)
        content.addLayout(date)
        content.addWidget(self.progress)
        content.addWidget(self.table)

    def dark_mode(self):
        self.setStyleSheet("""
            QWidget { background:#0f172a; color:#e5e7eb; font-size:14px; }
            QPushButton { background:#1e293b; padding:10px; border:none; }
            QPushButton:hover { background:#334155; }
            QTableWidget { background:#111827; gridline-color:#334155; }
            QHeaderView::section { background:#1f2937; color:#e5e7eb; padding:6px; }
        """)

    def set_table(self, headers, rows):
        self.table.clear()
        self.table.setRowCount(len(rows))
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)

        for r, row in enumerate(rows):
            for c, v in enumerate(row):
                self.table.setItem(r, c, QTableWidgetItem(str(v)))

        self.table.resizeColumnsToContents()

    def go_back(self):
        self.selected_toko = None
        self.mode = None
        self.title.setText("READY")
        self.table.clear()

    def load_data(self):
        folder = QFileDialog.getExistingDirectory()
        if not folder:
            return

        self.last_folder = folder
        self.start_loader(folder, manual=True)

    def start_loader(self, folder, manual=False):
        self.loader = Loader(folder)
        self.loader.progress.connect(
            lambda d, t: self.progress.setValue(int(d / t * 100)) if t else self.progress.setValue(0)
        )
        if manual:
            self.loader.finished.connect(
                lambda n: QMessageBox.information(self, "OK", f"{n} data baru")
            )
        else:
            self.loader.finished.connect(self.auto_reload_done)
        self.loader.start()

    def auto_reload(self):
        if not self.last_folder:
            return
        if hasattr(self, "loader") and self.loader.isRunning():
            return
        self.start_loader(self.last_folder, manual=False)

    def auto_reload_done(self, n):
        self.title.setText(f"AUTO REFRESH OK | {n} data baru")

    def rekap_tanggal(self):
        t1 = self.tgl1.date().toString("yyMMdd")
        t2 = self.tgl2.date().toString("yyMMdd")

        conn = sqlite3.connect(DB)
        rows = conn.execute("""
            SELECT tanggal, SUM(qty)
            FROM orders
            WHERE tanggal BETWEEN ? AND ?
            GROUP BY tanggal
            ORDER BY tanggal
        """, (t1, t2)).fetchall()
        conn.close()

        self.mode = "tanggal"
        self.set_table(
            ["No", "Tanggal", "Total Barang"],
            [(i + 1, r[0], r[1]) for i, r in enumerate(rows)]
        )
        self.title.setText("REKAP TOTAL BARANG PER TANGGAL")

    def rekap_toko(self):
        t1 = self.tgl1.date().toString("yyMMdd")
        t2 = self.tgl2.date().toString("yyMMdd")

        conn = sqlite3.connect(DB)
        rows = conn.execute("""
            SELECT toko, SUM(qty)
            FROM orders
            WHERE tanggal BETWEEN ? AND ?
            GROUP BY toko
            ORDER BY SUM(qty) DESC
        """, (t1, t2)).fetchall()
        conn.close()

        self.mode = "toko"
        self.set_table(
            ["No", "Toko", "Total Qty"],
            [(i + 1, *r) for i, r in enumerate(rows)]
        )
        self.title.setText("REKAP TOTAL BARANG PER TOKO")

    def rekap_semua_barang(self):
        t1 = self.tgl1.date().toString("yyMMdd")
        t2 = self.tgl2.date().toString("yyMMdd")

        conn = sqlite3.connect(DB)
        rows = conn.execute("""
            SELECT produk, SUM(qty)
            FROM orders
            WHERE tanggal BETWEEN ? AND ?
            GROUP BY produk
            ORDER BY SUM(qty) DESC, produk
        """, (t1, t2)).fetchall()
        conn.close()

        self.mode = "semua_barang"
        self.set_table(
            ["No", "Produk", "Total Qty", "Satuan"],
            [(i + 1, r[0], r[1], "pcs") for i, r in enumerate(rows)]
        )
        self.title.setText(f"REKAP SEMUA BARANG ({t1} s/d {t2})")

    def table_click(self, row, col):
        if self.mode == "tanggal":
            tanggal = self.table.item(row, 1).text()

            conn = sqlite3.connect(DB)
            rows = conn.execute("""
                SELECT produk, SUM(qty)
                FROM orders
                WHERE tanggal=?
                GROUP BY produk
                ORDER BY SUM(qty) DESC
            """, (tanggal,)).fetchall()
            conn.close()

            self.set_table(
                ["No", "Produk", "Total", "Satuan"],
                [(i + 1, r[0], r[1], "pcs") for i, r in enumerate(rows)]
            )
            self.title.setText(f"DETAIL PRODUK TANGGAL {tanggal}")
            return

        if self.mode == "toko":
            toko = self.table.item(row, 1).text()
            self.selected_toko = toko

            t1 = self.tgl1.date().toString("yyMMdd")
            t2 = self.tgl2.date().toString("yyMMdd")

            conn = sqlite3.connect(DB)
            rows = conn.execute("""
                SELECT produk, COUNT(DISTINCT oid), SUM(qty)
                FROM orders
                WHERE toko=? AND tanggal BETWEEN ? AND ?
                GROUP BY produk
                ORDER BY SUM(qty) DESC
            """, (toko, t1, t2)).fetchall()
            conn.close()

            self.set_table(
                ["No", "Produk", "TRX", "Qty"],
                [(i + 1, *r) for i, r in enumerate(rows)]
            )
            self.title.setText(f"DETAIL TOKO {toko}")

    def export_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save", "", "Excel (*.xlsx)")
        if not path:
            return

        rows = self.table.rowCount()
        cols = self.table.columnCount()
        if rows == 0 or cols == 0:
            QMessageBox.warning(self, "Kosong", "Tidak ada data untuk di-export")
            return

        headers = [self.table.horizontalHeaderItem(i).text() for i in range(cols)]
        data = [[self.table.item(r, c).text() for c in range(cols)] for r in range(rows)]
        pd.DataFrame(data, columns=headers).to_excel(path, index=False)

        QMessageBox.information(self, "OK", f"Export selesai:\n{path}")

    def export_excel_report_toko(self):
        if not self.selected_toko:
            QMessageBox.warning(self, "Pilih Toko", "Pilih toko dulu di Rekap Per Toko.")
            return

        t1 = self.tgl1.date().toString("yyMMdd")
        t2 = self.tgl2.date().toString("yyMMdd")

        path, _ = QFileDialog.getSaveFileName(self, "Save Report Toko", "", "Excel (*.xlsx)")
        if not path:
            return

        conn = sqlite3.connect(DB)
        rows = conn.execute("""
            SELECT o.tanggal,
                   o.toko,
                   o.produk,
                   SUM(o.qty) as total_qty,
                   CASE WHEN COALESCE(ps.is_taken, 0)=1 THEN 'SUDAH DIAMBIL' ELSE 'BELUM DIAMBIL' END as status,
                   COALESCE(ps.created_at, '-') as created_at,
                   COALESCE(ps.taken_at, '-') as taken_at
            FROM orders o
            LEFT JOIN product_status ps
              ON ps.toko = o.toko AND ps.tanggal = o.tanggal AND ps.produk = o.produk
            WHERE o.toko=? AND o.tanggal BETWEEN ? AND ?
            GROUP BY o.tanggal, o.toko, o.produk, ps.is_taken, ps.created_at, ps.taken_at
            ORDER BY o.tanggal, total_qty DESC, o.produk
        """, (self.selected_toko, t1, t2)).fetchall()
        conn.close()

        if not rows:
            QMessageBox.warning(self, "Kosong", "Tidak ada data report toko.")
            return

        df = pd.DataFrame(rows, columns=[
            "Tanggal Order",
            "Toko",
            "Produk",
            "Total Qty",
            "Status",
            "Tanggal Buat",
            "Tanggal Ambil"
        ])
        df.to_excel(path, index=False)

        QMessageBox.information(self, "OK", f"Report Excel selesai:\n{path}")

    def build_wa_message(self):
        rows = self.table.rowCount()
        if rows == 0:
            return None

        t1 = self.tgl1.date().toString("dd MMMM yyyy").upper()
        t2 = self.tgl2.date().toString("dd MMMM yyyy").upper()

        msg = "------------------------------------------\n"
        msg += f"REKAP TOKO {self.selected_toko}\n"
        msg += f"{t1} - {t2}\n\n"
        msg += "RINCIAN ORDER\n"
        msg += "------------------------------------------\n\n"

        total = 0
        for i in range(rows):
            produk = self.table.item(i, 1).text()
            qty = int(self.table.item(i, 3).text())
            total += qty
            msg += f"{i + 1}. {produk} ---> {qty} pcs\n"

        msg += "\n------------------------------------------\n"
        msg += f"TOTAL QTY ---> {total} pcs\n\n"
        msg += "Laporan otomatis ASP_GROUP\n"
        msg += "------------------------------------------"

        return msg

    def kirim_group(self):
        pesan = self.build_wa_message()
        if not pesan:
            QMessageBox.warning(self, "Kosong", "Tidak ada data")
            return

        QMessageBox.information(
            self,
            "Info",
            "Untuk versi HP, kirim report dilakukan lewat Telegram.\n"
            "Gunakan /menu7 lalu pilih toko dan klik 'Selesai & Kirim Report'."
        )


if __name__ == "__main__":
    init_db()

    bot_thread = threading.Thread(target=run_telegram_bot, daemon=True)
    bot_thread.start()

    app = QApplication([])
    win = App()
    win.show()
    app.exec()
