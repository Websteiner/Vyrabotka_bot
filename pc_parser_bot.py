#!/usr/bin/env python3
# pc_parser_bot.py — ПК-бот слушатель для сбора выручки.
# V4: 2026-11-16. Токены читаются из .env.telegram (безопасность).

import telebot
import openpyxl
import json
import requests
import os
import glob
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path
from dotenv import load_dotenv

# === ЗАГРУЗКА СЕКРЕТОВ ===
# Ищем .env.telegram сначала рядом со скриптом (для Windows .exe), потом в системной папке
SCRIPT_DIR = Path(__file__).resolve().parent
ENV_LOCAL = SCRIPT_DIR / ".env.telegram"
ENV_GLOBAL = Path.home() / ".config" / "bew" / "secrets" / ".env.telegram"

if ENV_LOCAL.exists():
    load_dotenv(ENV_LOCAL)
    print(f"[OK] Загружены секреты из {ENV_LOCAL}")
elif ENV_GLOBAL.exists():
    load_dotenv(ENV_GLOBAL)
    print(f"[OK] Загружены секреты из {ENV_GLOBAL}")
else:
    print("❌ Ошибка: файл .env.telegram не найден ни рядом со скриптом, ни в ~/.config/bew/secrets/")
    sys.exit(1)

PC_TOKEN = os.getenv("TELEGRAM_PC_TOKEN")
PHONE_TOKEN = os.getenv("TELEGRAM_PHONE_TOKEN")
CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

if not all([PC_TOKEN, PHONE_TOKEN, CHAT_ID]):
    print("❌ Ошибка: не все токены найдены в .env.telegram (нужны PC_TOKEN, PHONE_TOKEN, CHAT_ID)")
    sys.exit(1)

# === НАСТРОЙКИ ===
FILE_NAME_PATTERN = "*ИТОГОВЫЙ ОТЧЁТ ПО ВЫРУЧКЕ*.xlsx"

MONTHS_MAP = {
    "янв": "01", "фев": "02", "мар": "03", "апр": "04", "май": "05", "июн": "06",
    "июл": "07", "авг": "08", "сен": "09", "окт": "10", "ноя": "11", "дек": "12",
    "январь": "01", "февраль": "02", "март": "03", "апрель": "04", "май": "05", "июнь": "06",
    "июль": "07", "август": "08", "сентябрь": "09", "октябрь": "10", "ноябрь": "11", "декабрь": "12"
}
YEAR = "2026"

bot = telebot.TeleBot(PC_TOKEN)

def find_excel_file():
    user_home = os.path.expanduser("~")
    search_paths = [
        os.path.join(user_home, "Desktop"),
        os.path.join(user_home, "Downloads"),
        os.path.join(user_home, "Documents")
    ]
    found_files = []
    for path in search_paths:
        if os.path.exists(path):
            matches = glob.glob(os.path.join(path, FILE_NAME_PATTERN))
            for m in matches:
                # Исключаем временные файлы Excel (начинаются с ~$)
                if not os.path.basename(m).startswith("~$"):
                    found_files.append(m)
                    
    return max(found_files, key=os.path.getctime) if found_files else None

def parse_excel_mo_only(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    all_data = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        mo_row = None
        mo_col = None
        for row_idx, row in enumerate(ws.iter_rows(), start=1):
            for col_idx, cell in enumerate(row, start=1):
                if cell.value and "услуги мо" in str(cell.value).lower():
                    mo_row = row_idx
                    mo_col = col_idx
                    break
            if mo_row: break
        
        if not mo_row: continue
            
        header_row_idx = None
        col_date = None
        col_rev = None
        
        for check_row in range(mo_row, min(mo_row + 4, ws.max_row + 1)):
            date_c = None
            rev_c = None
            for col_idx in range(mo_col, min(mo_col + 20, ws.max_column + 1)):
                val = ws.cell(row=check_row, column=col_idx).value
                if val:
                    val_str = str(val).lower().strip()
                    if "дата" in val_str: date_c = col_idx
                    if "процедур" in val_str: rev_c = col_idx
            
            if date_c and rev_c:
                header_row_idx = check_row
                col_date = date_c
                col_rev = rev_c
                break
                
        if not header_row_idx: continue

        for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            if not row or len(row) < max(col_date, col_rev): continue
            
            date_val = row[col_date - 1]
            rev_val = row[col_rev - 1]
            
            if date_val and rev_val:
                try:
                    if isinstance(rev_val, str):
                        rev_val = float(rev_val.replace(" ", "").replace(",", ".").replace("₽", "").strip())
                    else:
                        rev_val = float(rev_val)
                except ValueError:
                    continue 

                date_str = None
                if isinstance(date_val, datetime):
                    date_str = date_val.strftime("%Y-%m-%d")
                elif isinstance(date_val, (int, float)):
                    date_str = (datetime(1899, 12, 30) + timedelta(days=date_val)).strftime("%Y-%m-%d")
                else:
                    date_raw = str(date_val).lower().strip()
                    match = re.search(r'(\d{1,2})[\.\s\-]?([а-яa-z]+)', date_raw)
                    if match:
                        day = match.group(1).zfill(2)
                        month_text = match.group(2)[:3]
                        month = MONTHS_MAP.get(month_text)
                        if month:
                            date_str = f"{YEAR}-{month}-{day}"

                if date_str and len(date_str) == 10:
                    all_data[date_str] = all_data.get(date_str, 0) + rev_val

    return all_data

def send_to_phone(message, data):
    json_path = os.path.join(os.environ.get('TEMP', '.'), "vyruchka_data.json")
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    url = f"https://api.telegram.org/bot{PHONE_TOKEN}/sendDocument"
    with open(json_path, 'rb') as f:
        files = {'document': f}
        payload = {'chat_id': CHAT_ID, 'caption': '📦 Данные о выручке (Услуги МО)'}
        response = requests.post(url, data=payload, files=files, timeout=15)
    
    if response.status_code == 200:
        bot.reply_to(message, "✅ Данные переданы на телефон!")
    else:
        bot.reply_to(message, f"❌ Ошибка отправки: {response.text}")

@bot.message_handler(commands=['collect'])
def collect_data(message):
    bot.reply_to(message, "⏳ Принял команду. Ищу файл...")
    excel_path = find_excel_file()
    if not excel_path:
        bot.reply_to(message, f"❌ Ошибка: Файл '{FILE_NAME_PATTERN}' не найден на Рабочем столе, в Загрузках или Документах.")
        return

    bot.reply_to(message, f"📄 Найден: {os.path.basename(excel_path)}\nПарсим...")

    try:
        data = parse_excel_mo_only(excel_path)
        if not data:
            bot.reply_to(message, "⚠️ Данные не найдены. Убедитесь, что есть блок 'Услуги МО' с колонками 'Дата' и 'Процедуры'.")
            return
        send_to_phone(message, data)
        
    except PermissionError:
        bot.reply_to(message, "⚠️ Файл заблокирован (открыт в Excel). Пожалуйста, закройте его на компьютере и попробуйте снова.")
    except Exception as e:
        bot.reply_to(message, f"❌ Ошибка парсинга: {e}")

if __name__ == '__main__':
    print("ПК-парсер запущен (Только Услуги МО, токены из .env)...")
    bot.polling(non_stop=True)
