import telebot
import pandas as pd
import os
from datetime import datetime, timedelta
import logging

# ==================== KONFIGURASI ====================
BOT_TOKEN = '8901282411:AAGMR2yxbrfEtf_DvX-lrZ_Wd6MHqnRpJ9s'  # Ganti jika perlu
EXCEL_FILE = 'user_data.xlsx'

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
bot = telebot.TeleBot(BOT_TOKEN)

bot.set_my_commands([
    telebot.types.BotCommand("start", "Memulai dan menampilkan bantuan bot"),
    telebot.types.BotCommand("help", "Menampilkan daftar perintah"),
    telebot.types.BotCommand("lihat", "Menampilkan 10 pesan terakhir yang tersimpan"),
    telebot.types.BotCommand("download", "Mengirim file Excel berisi data"),
    telebot.types.BotCommand("hapus", "Menghapus semua data (butuh konfirmasi)"),
    telebot.types.BotCommand("hapus_pesan", "Menghapus pesan tertentu (pilih dari daftar)"),
    telebot.types.BotCommand("batal", "Membatalkan proses penghapusan"),
])

# ==================== FUNGSI DASAR ====================
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=['User ID', 'Username', 'Message', 'Date'])
        df.to_excel(EXCEL_FILE, index=False)
        logging.info("File Excel baru dibuat.")

def is_duplicate(user_id, message_text):
    try:
        df = pd.read_excel(EXCEL_FILE)
        if df.empty:
            return False
        user_df = df[df['User ID'] == user_id]
        if user_df.empty:
            return False
        last_row = user_df.iloc[-1]
        last_message = last_row['Message']
        last_time = last_row['Date']
        if last_message != message_text:
            return False
        time_diff = datetime.now() - last_time
        return time_diff < timedelta(hours=1)
    except Exception as e:
        logging.error(f"Error cek duplikat: {e}")
        return False

def save_message_to_excel(user_id, username, message_text):
    try:
        if is_duplicate(user_id, message_text):
            return False, "Anda telah mengirim pesan yang sama dalam 1 jam terakhir. Pesan tidak disimpan."
        new_data = pd.DataFrame([{
            'User ID': user_id,
            'Username': username,
            'Message': message_text,
            'Date': datetime.now()
        }])
        if not os.path.exists(EXCEL_FILE):
            new_data.to_excel(EXCEL_FILE, index=False)
        else:
            existing_df = pd.read_excel(EXCEL_FILE)
            combined_df = pd.concat([existing_df, new_data], ignore_index=True)
            combined_df.to_excel(EXCEL_FILE, index=False)
        return True, "✅ Pesan berhasil disimpan!"
    except Exception as e:
        logging.error(f"Gagal simpan: {e}")
        return False, f"❌ Gagal menyimpan: {e}"

def get_all_data_as_text(limit=10):
    try:
        df = pd.read_excel(EXCEL_FILE)
        if df.empty:
            return "📭 Belum ada data tersimpan."
        display_df = df.tail(limit)
        result = f"📊 *{len(df)} pesan total* (menampilkan {len(display_df)} terakhir):\n\n"
        for idx, row in display_df.iterrows():
            nomor = idx + 1
            username = row['Username']
            msg = row['Message'][:40] + "..." if len(str(row['Message'])) > 40 else row['Message']
            date = row['Date'].strftime("%d/%m %H:%M") if hasattr(row['Date'], 'strftime') else str(row['Date'])[:16]
            result += f"{nomor}. @{username}: \"{msg}\" ({date})\n"
        return result
    except Exception as e:
        return f"❌ Error membaca data: {e}"

def delete_all_data():
    try:
        empty_df = pd.DataFrame(columns=['User ID', 'Username', 'Message', 'Date'])
        empty_df.to_excel(EXCEL_FILE, index=False)
        return True, "🗑️ Semua data berhasil dihapus!"
    except Exception as e:
        return False, f"❌ Gagal menghapus data: {e}"

# ==================== FITUR HAPUS PESAN TERTENTU ====================
def get_user_messages_with_index(user_id, limit=10):
    try:
        df = pd.read_excel(EXCEL_FILE)
        if df.empty:
            return None, "Belum ada data."
        user_df = df[df['User ID'] == user_id].copy()
        if user_df.empty:
            return None, "Anda belum memiliki pesan tersimpan."
        user_df = user_df.sort_values('Date', ascending=False)
        user_df['original_index'] = user_df.index
        user_df = user_df.reset_index(drop=True)
        user_df.insert(0, 'No', range(1, len(user_df)+1))
        display_df = user_df.head(limit)
        return display_df, None
    except Exception as e:
        return None, f"Error: {e}"

def delete_message_by_original_index(original_index):
    try:
        df = pd.read_excel(EXCEL_FILE)
        if original_index not in df.index:
            return False, "Pesan tidak ditemukan."
        df = df.drop(original_index)
        df = df.reset_index(drop=True)
        df.to_excel(EXCEL_FILE, index=False)
        return True, "✅ Pesan berhasil dihapus."
    except Exception as e:
        return False, f"❌ Gagal menghapus: {e}"

user_selection_data = {}

# ==================== HANDLER PERINTAH ====================
@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.reply_to(message, 
        "🤖 *Bot Pencatat Pesan*\n\n"
        "Kirim pesan apa saja, akan saya simpan ke Excel.\n\n"
        "*Perintah yang tersedia:*\n"
        "/lihat - Menampilkan 10 pesan terakhir semua user\n"
        "/download - Mengirim file Excel data pesan\n"
        "/hapus - Menghapus SEMUA data (perlu konfirmasi YA)\n"
        "/hapus_pesan - Menghapus pesan tertentu milik Anda\n"
        "/batal - Membatalkan proses penghapusan\n"
        "/help - Menampilkan bantuan ini",
        parse_mode='Markdown')

@bot.message_handler(commands=['lihat'])
def lihat_data(message):
    text = get_all_data_as_text(10)
    bot.reply_to(message, text, parse_mode='Markdown')

@bot.message_handler(commands=['download'])
def download_excel(message):
    try:
        if not os.path.exists(EXCEL_FILE):
            bot.reply_to(message, "❌ Belum ada data. File Excel belum dibuat.")
            return
        with open(EXCEL_FILE, 'rb') as f:
            bot.send_document(message.chat.id, f, caption="📊 Berikut file Excel berisi data pesan.")
        logging.info(f"User {message.from_user.id} mendownload file Excel.")
    except Exception as e:
        logging.error(f"Gagal mengirim file: {e}")
        bot.reply_to(message, f"❌ Gagal mengirim file: {e}")

user_waiting_confirm = set()

@bot.message_handler(commands=['hapus'])
def delete_confirm(message):
    user_waiting_confirm.add(message.chat.id)
    bot.reply_to(message, 
        "⚠️ *PERINGATAN!*\n\n"
        "Anda akan menghapus SEMUA data yang tersimpan.\n"
        "Ketik `/batal` untuk membatalkan, atau kirim `YA` untuk konfirmasi.",
        parse_mode='Markdown')

@bot.message_handler(commands=['batal'])
def cancel_delete(message):
    chat_id = message.chat.id
    if chat_id in user_waiting_confirm:
        user_waiting_confirm.remove(chat_id)
    if chat_id in user_selection_data:
        del user_selection_data[chat_id]
    bot.reply_to(message, "Penghapusan dibatalkan.")

@bot.message_handler(func=lambda message: message.text and message.text.upper() == "YA")
def execute_delete(message):
    if message.chat.id in user_waiting_confirm:
        user_waiting_confirm.remove(message.chat.id)
        success, msg = delete_all_data()
        bot.reply_to(message, msg)
    else:
        bot.reply_to(message, "Tidak ada konfirmasi penghapusan. Gunakan /hapus dulu.")

@bot.message_handler(commands=['hapus_pesan'])
def cmd_hapus_pesan(message):
    user_id = message.from_user.id
    chat_id = message.chat.id
    display_df, err = get_user_messages_with_index(user_id)
    if err:
        bot.reply_to(message, err)
        return
    if display_df is None or display_df.empty:
        bot.reply_to(message, "Tidak ada pesan yang bisa dihapus.")
        return
    text = "📋 *Pilih nomor pesan yang ingin dihapus:*\n\n"
    for _, row in display_df.iterrows():
        no = row['No']
        msg = row['Message'][:50] + "..." if len(str(row['Message'])) > 50 else row['Message']
        date = row['Date'].strftime("%d/%m %H:%M") if hasattr(row['Date'], 'strftime') else str(row['Date'])[:16]
        text += f"{no}. {msg} ({date})\n"
    text += f"\nKetik angka (1-{len(display_df)}) untuk menghapus. Ketik /batal untuk membatalkan."
    sent_msg = bot.reply_to(message, text, parse_mode='Markdown')
    user_selection_data[chat_id] = {'display_df': display_df, 'prompt_msg_id': sent_msg.message_id}

@bot.message_handler(func=lambda message: message.chat.id in user_selection_data and message.text.isdigit())
def process_delete_choice(message):
    chat_id = message.chat.id
    choice = int(message.text)
    data = user_selection_data.get(chat_id)
    if not data:
        return
    display_df = data['display_df']
    if choice < 1 or choice > len(display_df):
        bot.reply_to(message, "Nomor tidak valid. Silakan coba lagi.")
        return
    selected_row = display_df[display_df['No'] == choice].iloc[0]
    original_index = selected_row['original_index']
    success, msg = delete_message_by_original_index(original_index)
    bot.reply_to(message, msg)
    del user_selection_data[chat_id]

@bot.message_handler(func=lambda message: True)
def handle_message(message):
    if message.text.startswith('/'):
        return
    user_id = message.from_user.id
    username = message.from_user.username or message.from_user.first_name
    user_message = message.text
    success, reply_msg = save_message_to_excel(user_id, username, user_message)
    bot.reply_to(message, reply_msg)

# ==================== JALANKAN BOT ====================
if __name__ == '__main__':
    init_excel()
    logging.info("🚀 Bot sedang berjalan...")
    try:
        bot.infinity_polling(timeout=10, long_polling_timeout=5)
    except Exception as e:
        logging.error(f"Bot berhenti: {e}")