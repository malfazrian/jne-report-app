import email
import re
import os
import subprocess
import time
import pyautogui
import py7zr
from email.policy import default
from email.utils import parsedate_to_datetime
from datetime import datetime
from bs4 import BeautifulSoup
from typing import Optional
from zipfile import BadZipFile
from workflows.file_ops import extract_archive_with_password

def buka_thunderbird():
    subprocess.Popen(["C:\\Program Files\\Mozilla Thunderbird Beta\\thunderbird.exe"])
    print("Thunderbird dibuka...")
    time.sleep(10)  # tunggu sampai aplikasi terbuka

def refresh_inbox(mbox_path, timeout=300):
    last_mtime = os.path.getmtime(mbox_path)
    pyautogui.press('f5')
    print("Memulai proses refresh...")

    start_time = time.time()
    while True:
        time.sleep(2)
        new_mtime = os.path.getmtime(mbox_path)
        if new_mtime != last_mtime:
            print("Refresh selesai.")
            break
        if time.time() - start_time > timeout:
            print("Batas waktu 5 menit tercapai, proses menunggu refresh dihentikan..")
            break

def parse_custom_date(date_str):
    try:
        return datetime.strptime(date_str, "%d-%b-%Y %H:%M:%S")
    except Exception:
        return None
    
def normalize_subject(subject: str) -> str:
    return subject.lower().replace("re:", "").replace("fwd:", "").strip()

def find_latest_matching_email(file_path, subject_prefix, from_email: Optional[str] = None, max_emails=100):
    latest_email = None
    latest_date = None
    email_count = 0
    current_block = []
    blocks = []

    with open(file_path, encoding="utf-8", errors="ignore") as f:
        for line in f:
            if line.startswith("X-Account-Key:"):
                if current_block:
                    blocks.append("".join(current_block))
                    if len(blocks) > max_emails:
                        blocks.pop(0)
                    current_block = []
            current_block.append(line)
        if current_block:
            blocks.append("".join(current_block))
            if len(blocks) > max_emails:
                blocks.pop(0)

    for raw in reversed(blocks):
        msg = email.message_from_string(raw, policy=default)
        email_count += 1
        if email_count % 10 == 0:
            print(f"Sudah baca {email_count} email dari belakang...")

        from_ = msg.get("From", "")
        subject = msg.get("Subject", "")
        date_str = msg.get("Date")
        try:
            date = parsedate_to_datetime(date_str) if date_str else None
        except Exception:
            date = None

        # Coba parse custom jika gagal
        if not date and date_str:
            date = parse_custom_date(date_str)
            if not date:
                print("Gagal parse tanggal:", date_str)

        # Cek kecocokan subject dan (jika ada) from_email
        subject_match = subject and subject_prefix.lower() in normalize_subject(subject)
        from_match = True if not from_email else from_email.lower() in from_.lower()

        if subject_match and from_match:
            if date and (latest_date is None or date > latest_date):
                latest_email = msg
                latest_date = date

    return latest_email

def extract_login_info_html(html_body: str):
    soup = BeautifulSoup(html_body, "html.parser")

    # Ambil link dari tag <a>
    link = None
    a_tag = soup.find("a", href=True)
    if a_tag:
        link = a_tag['href']

    # Cari username dan password dari teks
    username = None
    password = None
    text = soup.get_text(separator="\n")
    for line in text.splitlines():
        if "username:" in line.lower():
            username = line.split(":", 1)[-1].strip()
        elif "password:" in line.lower():
            password = line.split(":", 1)[-1].strip()

    if link and username and password:
        return {
            "Link": link,
            "Username": username,
            "Password": password
        }
    return None

# Ambil body dari email
def get_email_body(msg):
    if msg.is_multipart():
        # Coba ambil text/plain dulu
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                return part.get_payload(decode=True).decode(errors="ignore")
        # Kalau gagal, ambil text/html
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                return part.get_payload(decode=True).decode(errors="ignore")
    else:
        return msg.get_payload(decode=True).decode(errors="ignore")
    return ""

def save_attachments(msg, save_dir="attachments"):
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    count = 0
    for part in msg.walk():
        # Skip if not attachment
        if part.get_content_disposition() != 'attachment':
            continue

        filename = part.get_filename()
        if filename:
            filepath = os.path.join(save_dir, filename)
            with open(filepath, "wb") as f:
                f.write(part.get_payload(decode=True))
            print(f"Attachment disimpan: {filepath}")
            count += 1

            # Ekstrak jika .7z (tanpa password)
            if filename.endswith(".7z"):
                try:
                    with py7zr.SevenZipFile(filepath, mode='r') as archive:
                        archive.extractall(path=save_dir)
                        print(f"Berhasil extract: {filename}")
                except Exception as e:
                    print(f"Gagal extract {filename}: {e}")

    if count == 0:
        print("Tidak ada attachment ditemukan.")
    else:
        print(f"Total attachment disimpan: {count}")

def save_attachments_danamon(msg, save_dir="attachments"):
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    subject = msg.get("Subject", "")
    print("Subject email:", subject)

    password = None

    # --- 1. Coba format angka: dd/mm/yy atau dd-mm-yy
    match = re.search(r"(\d{2})[/-]\d{2}[/-]\d{2}", subject)
    if match:
        dd = match.group(1)
        password = f"danamon{dd}"
    else:
        # --- 2. Coba format nama bulan Indonesia: 14 JULI 2025
        match_bulan = re.search(r"(\d{1,2})\s+([A-Z]+)\s+(\d{4})", subject.upper())
        bulan_mapping = {
            "JANUARI": 1, "FEBRUARI": 2, "MARET": 3, "APRIL": 4,
            "MEI": 5, "JUNI": 6, "JULI": 7, "AGUSTUS": 8, "SEPTEMBER": 9,
            "OKTOBER": 10, "NOVEMBER": 11, "DESEMBER": 12
        }
        if match_bulan:
            dd = match_bulan.group(1).zfill(2)  # Pastikan 2 digit
            bulan_nama = match_bulan.group(2)
            if bulan_nama in bulan_mapping:
                password = f"danamon{dd}"
            else:
                print("Nama bulan tidak dikenali:", bulan_nama)

    if not password:
        print("Tidak ditemukan tanggal dalam subject. ZIP/7Z tidak akan diextract.")

    count = 0
    for part in msg.walk():
        if part.get_content_disposition() != 'attachment':
            continue

        filename = part.get_filename()
        if not filename:
            continue

        filepath = os.path.join(save_dir, filename)
        with open(filepath, "wb") as f:
            f.write(part.get_payload(decode=True))
        print(f"Attachment disimpan: {filepath}")
        count += 1

        # Ekstrak jika ZIP/7Z dan password tersedia
        if filename.lower().endswith((".zip", ".7z")) and password:
            extract_archive_with_password(
                file_path=filepath,
                extract_to=save_dir,
                password=password,
                delete_after=True
            )

    if count == 0:
        print("Tidak ada attachment ditemukan.")
    else:
        print(f"Total attachment disimpan: {count}")

def extract_date_from_subject(subject: str) -> Optional[str]:
    """
    Ekstrak tanggal dari subject email dan konversi ke format DD/MM/YYYY.

    Contoh subject:
        - "DATA PICKUP JNE 30/06/25" → "2025-06-30"
        - "OS CARD JNE 01/07/2025"   → "2025-07-01"

    Returns:
        str (format DD/MM/YYYY) jika berhasil, None jika gagal.
    """
    if not subject:
        return None

    # Cari tanggal dengan format DD/MM/YY atau DD/MM/YYYY
    match = re.search(r"(\d{2})/(\d{2})/(\d{2,4})", subject)
    if not match:
        return None

    day, month, year = match.groups()
    if len(year) == 2:
        year = "20" + year  # Anggap tahun 2000-an

    try:
        date_obj = datetime.strptime(f"{day}/{month}/{year}", "%d/%m/%Y")
        return date_obj.strftime("%d/%m/%y")
    except ValueError:
        return None