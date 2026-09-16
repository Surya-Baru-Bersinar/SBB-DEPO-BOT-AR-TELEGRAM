import configparser
import io
import logging
import os
import re
import threading
import time
from datetime import datetime

import matplotlib
import pandas as pd
import telebot
from telebot import types

matplotlib.use('Agg')
import matplotlib.pyplot as plt

telebot.logger.setLevel(logging.CRITICAL)

session_lock = threading.Lock()
data_lock = threading.Lock()

authenticated_users = set()
user_sessions = {}
G_DATA = {
    'df_ar': None,
    'last_updated': None
}

def load_config(config_file='config.conf'):
    if not os.path.exists(config_file):
        raise ValueError(f"CRITICAL ERROR: File konfigurasi '{config_file}' tidak ditemukan!")
    
    config = configparser.ConfigParser()
    config.read(config_file, encoding='utf-8')

    secret_key = config.get('AUTH', 'secret_key', fallback='').strip()
    bot_token = config.get('BOT', 'bot_token', fallback='').strip()

    if not secret_key:
        raise ValueError("CRITICAL ERROR: 'secret_key' di seksi [AUTH] wajib diisi!")
    if not bot_token:
        raise ValueError("CRITICAL ERROR: 'bot_token' di seksi [BOT] wajib diisi!")

    return config

config = load_config()
SECRET_KEY = config.get('AUTH', 'secret_key').strip()
BOT_TOKEN = config.get('BOT', 'bot_token').strip()
COMPANY_NAME = config.get('GENERAL', 'gen_company_name', fallback='PT ABC').strip()
DEPO_CONFIG = config.get('MAP', 'depo', fallback='SL|YY|MKS|MGL|PW|PWT|PLU|SG|SMG|TGL|PA|KDI').strip()
AR_FILE_PATH = config.get('DIR', 'arvi', fallback='ARClean_temp.xlsx').strip()

bot = telebot.TeleBot(BOT_TOKEN)

def standardize_code(code, depo_prefixes=DEPO_CONFIG):
    if pd.isna(code):
        return ""
    if isinstance(code, float) and code.is_integer():
        code = int(code)
    s = str(code).strip().upper()
    s = re.sub(r'\s*-\s*', '-', s)
    pattern = rf'^({depo_prefixes})\s*(\d+)'
    s = re.sub(pattern, r'\1-\2', s)
    s = re.sub(r'(\d+)([A-Z])$', r'\1 \2', s)
    return s

def potong_teks(teks, max_char=35):
    if pd.isna(teks):
        return ""
    s = str(teks).strip()
    return s[:max_char-3] + '...' if len(s) > max_char else s

def format_tanggal_jt(teks):
    if pd.isna(teks):
        return ""
    s = str(teks).strip()
    if s.lower() in ['nan', 'none', 'nat', '']:
        return ""
    if '&' in s:
        parts = [p.strip() for p in s.split('&') if p.strip()]
        return " &\n".join(parts)
    return s

def is_tanggal_jt_terisi(val):
    if pd.isna(val):
        return False
    s = str(val).strip().lower()
    return s not in ['', 'nan', 'none', 'nat']

def read_excel_auto_header(file_path, target_column="Kode Pelanggan"):
    df_raw = pd.read_excel(file_path, sheet_name=0, header=None)
    target_clean = str(target_column).strip().upper()

    for idx, row in df_raw.iterrows():
        row_cleaned = [str(val).strip().upper() for val in row.dropna()]
        if target_clean in row_cleaned or "NO. PELANGGAN" in row_cleaned or "NAMA PELANGGAN" in row_cleaned:
            df_clean = df_raw.iloc[idx + 1 :].copy()
            df_clean.columns = df_raw.iloc[idx].astype(str).str.strip()
            return df_clean.reset_index(drop=True)

    df_raw.columns = df_raw.iloc[0].astype(str).str.strip()
    return df_raw.iloc[1:].reset_index(drop=True)

def load_ar_dataset_from_disk():
    if not os.path.exists(AR_FILE_PATH):
        raise FileNotFoundError(f"File data AR '{AR_FILE_PATH}' tidak ditemukan!")

    df_ar = read_excel_auto_header(AR_FILE_PATH, target_column="Kode Pelanggan")

    def clean_raw_code(val):
        if pd.isna(val):
            return ""
        if isinstance(val, (int, float)):
            return str(int(val))
        s = str(val).strip()
        if s.endswith('.0'):
            s = s[:-2]
        return s.upper()

    col_nopel = None
    for col in df_ar.columns:
        if str(col).strip().lower() in ['kode pelanggan', 'no. pelanggan', 'nopel', 'no pelanggan']:
            col_nopel = col
            break

    if col_nopel:
        df_ar['Raw_Kode'] = df_ar[col_nopel].apply(clean_raw_code)
        df_ar['Clean_Kode'] = df_ar[col_nopel].apply(lambda x: standardize_code(x, DEPO_CONFIG))
    else:
        df_ar['Raw_Kode'] = ""
        df_ar['Clean_Kode'] = ""

    return df_ar

def background_data_refresher():
    while True:
        try:
            df_ar = load_ar_dataset_from_disk()
            with data_lock:
                G_DATA['df_ar'] = df_ar
                G_DATA['last_updated'] = datetime.now()
            print(f"--> [{datetime.now().strftime('%H:%M:%S')}] Caching Data AR Berhasil!")
        except Exception as err:
            print(f"--> [ERROR REFRESH DATA]: {err}")

        time.sleep(600)

def generate_ar_image(df_filtered):
    df = df_filtered.copy()

    if 'No. Faktur' in df.columns:
        df = df.sort_values(by='No. Faktur', ascending=True).reset_index(drop=True)

    if df.empty:
        return None

    cols_to_show = [
        'No. Faktur', 'Tgl Faktur', 'Jatuh Tempo', 'Nilai Faktur', 
        'Sisa Piutang', 'Umur JT', 'Nama Pelanggan', 'Nama Penjual', 
        'Nama Kontak'
    ]

    if 'Tanggal JT' in df.columns:
        cols_to_show.append('Tanggal JT')

    valid_cols = [c for c in cols_to_show if c in df.columns]
    df_display = df[valid_cols].copy()

    if 'Nama Kontak' in df_display.columns:
        df_display['Nama Kontak'] = df_display['Nama Kontak'].apply(lambda x: potong_teks(x, max_char=35))
    if 'Nama Pelanggan' in df_display.columns:
        df_display['Nama Pelanggan'] = df_display['Nama Pelanggan'].apply(lambda x: potong_teks(x, max_char=30))
    if 'Nama Penjual' in df_display.columns:
        df_display['Nama Penjual'] = df_display['Nama Penjual'].apply(lambda x: potong_teks(x, max_char=16))

    if 'Tanggal JT' in df_display.columns:
        df_display['Tanggal JT'] = df_display['Tanggal JT'].apply(format_tanggal_jt)

    total_nilai = 0
    total_sisa = 0

    for col in ['Nilai Faktur', 'Sisa Piutang']:
        if col in df_display.columns:
            numeric_vals = pd.to_numeric(df_display[col], errors='coerce').fillna(0)
            if col == 'Nilai Faktur':
                total_nilai = numeric_vals.sum()
            elif col == 'Sisa Piutang':
                total_sisa = numeric_vals.sum()
            
            df_display[col] = numeric_vals.apply(lambda x: f"{x:,.0f}".replace(",", "."))

    num_rows = len(df_display)
    width_inches = 22.0
    height_inches = max(2.5, num_rows * 0.5 + 1.2)
    
    fig, ax = plt.subplots(figsize=(width_inches, height_inches))
    ax.axis('off')

    navy_color = (0.0, 0.125, 0.376)

    table = ax.table(
        cellText=df_display.values, 
        colLabels=df_display.columns, 
        bbox=[0.0, 0.11, 1.0, 0.81],
        cellLoc='left'
    )
    table.auto_set_font_size(False)
    table.set_fontsize(9.5)

    for (row_idx, _), cell in table.get_celld().items():
        if row_idx == 0:
            cell.set_facecolor(navy_color)
            cell.set_text_props(color='white', weight='bold', verticalalignment='center')
        else:
            cell.set_text_props(verticalalignment='center')

    plt.figtext(0.5, 0.94, COMPANY_NAME, ha="center", va="bottom", fontsize=15, fontweight='bold', color=navy_color)

    summary_title = f"Total Nilai Faktur: Rp {total_nilai:,.0f}  |  Total Sisa Piutang: Rp {total_sisa:,.0f}".replace(",", ".")
    plt.figtext(0.5, 0.04, summary_title, ha="center", va="top", fontsize=11, fontweight='bold', color=navy_color)

    img_buffer = io.BytesIO()
    plt.savefig(img_buffer, format='png', bbox_inches='tight', pad_inches=0.2, dpi=200)
    plt.close(fig)
    img_buffer.seek(0)

    return img_buffer

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    user_id = message.from_user.id
    with session_lock:
        is_auth = user_id in authenticated_users

    if not is_auth:
        bot.reply_to(message, "Akses Terkunci.\nMasukkan Sandi/Token Internal untuk mengaktifkan bot:")
    else:
        bot.reply_to(message, "Bot AR Ready!\nSilakan masukkan Nomor Pelanggan (Nopel).\nContoh: 7159 atau YY-2223 & SL-1122")

@bot.message_handler(func=lambda msg: True)
def handle_incoming_messages(message):
    user_id = message.from_user.id

    with session_lock:
        is_auth = user_id in authenticated_users

    if not is_auth:
        if message.text.strip() == SECRET_KEY:
            with session_lock:
                authenticated_users.add(user_id)
            bot.reply_to(message, "Akses Diterima! Silakan masukkan Nomor Pelanggan (Nopel).")
        else:
            bot.reply_to(message, "Sandi Salah! Silakan coba lagi.")
        return

    query = message.text.strip()

    with data_lock:
        df_ar = G_DATA['df_ar']

    if df_ar is None or df_ar.empty:
        bot.reply_to(message, "Data AR belum siap di memori RAM. Silakan tunggu beberapa saat.")
        return

    clean_q = query.lower().replace('nopel:', '').strip()
    raw_codes = [k.strip() for k in clean_q.split('&') if k.strip()]

    target_codes_set = set()
    for r_code in raw_codes:
        code_upper = r_code.upper()
        std_c = standardize_code(r_code, DEPO_CONFIG)
        target_codes_set.add(code_upper)
        target_codes_set.add(std_c)

    cond_raw = df_ar['Raw_Kode'].isin(target_codes_set)
    cond_clean = df_ar['Clean_Kode'].isin(target_codes_set)
    matched_df = df_ar[cond_raw | cond_clean]

    if matched_df.empty:
        bot.reply_to(message, f"Data piutang tidak ditemukan untuk Nomor Pelanggan: '{query}'.")
        return

    has_tgl_jt = 'Tanggal JT' in matched_df.columns

    if has_tgl_jt:
        with session_lock:
            user_sessions[user_id] = {
                'data': matched_df,
                'query': query
            }

        markup = types.InlineKeyboardMarkup(row_width=2)
        markup.add(
            types.InlineKeyboardButton("Tampilkan Tanggal JT", callback_data="tgljt_SHOW"),
            types.InlineKeyboardButton("Sembunyikan Data Tanggal JT", callback_data="tgljt_HIDE")
        )
        bot.send_message(
            message.chat.id, 
            f"Ditemukan {len(matched_df)} faktur piutang.\n\nPilih opsi penanganan Tanggal JT:", 
            reply_markup=markup
        )
    else:
        bot.send_message(message.chat.id, f"Ditemukan {len(matched_df)} faktur piutang. Mengolah gambar laporan...")
        img_buffer = generate_ar_image(matched_df)

        if img_buffer:
            caption_msg = f"Laporan Piutang AR - Nopel: {query}"
            try:
                bot.send_photo(message.chat.id, img_buffer, caption=caption_msg)
            except Exception:
                img_buffer.seek(0)
                img_buffer.name = f"Laporan_Piutang_{query}.png"
                bot.send_document(message.chat.id, img_buffer, caption=caption_msg)

@bot.callback_query_handler(func=lambda call: call.data.startswith('tgljt_'))
def process_tgljt_callback(call):
    user_id = call.from_user.id
    
    with session_lock:
        if user_id not in user_sessions:
            bot.answer_callback_query(call.id, "Sesi pencarian telah berakhir. Silakan masukkan Nopel kembali.")
            return
        session_data = user_sessions[user_id]

    df_data = session_data['data'].copy()
    query_nopel = session_data['query']
    action = call.data

    if action == "tgljt_HIDE":
        df_data = df_data[~df_data['Tanggal JT'].apply(is_tanggal_jt_terisi)].reset_index(drop=True)
        status_text = "Data Ber-Tanggal JT Disembunyikan"
    else:
        status_text = "Semua Data Ditampilkan (Termasuk Tanggal JT)"

    try:
        bot.edit_message_text(
            f"Memproses gambar laporan ({status_text})...", 
            chat_id=call.message.chat.id, 
            message_id=call.message.message_id
        )
    except Exception:
        pass

    if df_data.empty:
        bot.send_message(call.message.chat.id, "Tidak ada data piutang tersisa setelah disaring.")
        return

    img_buffer = generate_ar_image(df_data)

    if img_buffer:
        caption_msg = f"Laporan Piutang AR - Nopel: {query_nopel}\nStatus: {status_text}"
        try:
            bot.send_photo(call.message.chat.id, img_buffer, caption=caption_msg)
        except Exception:
            img_buffer.seek(0)
            img_buffer.name = f"Laporan_Piutang_{query_nopel}.png"
            bot.send_document(call.message.chat.id, img_buffer, caption=caption_msg)
    else:
        bot.send_message(call.message.chat.id, "Gagal memproses gambar laporan.")

if __name__ == '__main__':
    print("--> Memulai Data RAM Preloader...")
    try:
        df_ar_init = load_ar_dataset_from_disk()
        with data_lock:
            G_DATA['df_ar'] = df_ar_init
            G_DATA['last_updated'] = datetime.now()
        print("--> Data AR awal berhasil dimuat!")
    except Exception as e:
        print(f"--> [WARNING INITIATION]: {e}")

    refresher_thread = threading.Thread(target=background_data_refresher, daemon=True)
    refresher_thread.start()

    print("--> Bot Telegram Piutang AR Aktif...")

    while True:
        try:
            bot.infinity_polling(timeout=20, long_polling_timeout=15)
        except Exception as err:
            print(f"--> [ERROR POLLING]: {err}")
            time.sleep(5)