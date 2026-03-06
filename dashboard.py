from telegram import InlineKeyboardButton, InlineKeyboardMarkup


def dashboard():

    return InlineKeyboardMarkup([

        [
            InlineKeyboardButton("📥 Upload Order TXT",callback_data="menu|upload"),
            InlineKeyboardButton("🔎 Search Order",callback_data="menu|search"),
        ],

        [
            InlineKeyboardButton("📦 Rekap Produk",callback_data="menu|produk"),
            InlineKeyboardButton("🏪 Rekap Toko",callback_data="menu|toko"),
        ],

        [
            InlineKeyboardButton("📊 Export Excel",callback_data="menu|excel"),
            InlineKeyboardButton("📡 Status Server",callback_data="menu|status"),
        ],

        [
            InlineKeyboardButton("❓ Bantuan",callback_data="menu|help"),
        ]
    ])