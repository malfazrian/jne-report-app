import email
import re
import os
import subprocess
import time
import pyautogui
from email.policy import default
from email.utils import parsedate_to_datetime
from datetime import datetime
from bs4 import BeautifulSoup
from typing import Optional
from zipfile import BadZipFile
from workflows.file_ops import extract_zip_with_password

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
        subject_match = subject and subject.startswith(subject_prefix)
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
            print(f"📎 Attachment disimpan: {filepath}")
            count += 1

    if count == 0:
        print("Tidak ada attachment ditemukan.")
    else:
        print(f"Total attachment disimpan: {count}")

def save_attachments_danamon(msg, save_dir="attachments"):
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    subject = msg.get("Subject", "")
    print("Subject email:", subject)

    # Cari tanggal dari subject (format 30/06/25)
    match = re.search(r"(\d{2})/\d{2}/\d{2}", subject)
    if match:
        dd = match.group(1)
        password = f"danamon{dd}"
    else:
        print("Tidak ditemukan tanggal dalam subject. ZIP tidak akan diextract.")
        password = None

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
        print(f"📎 Attachment disimpan: {filepath}")
        count += 1

        # Jika zip, langsung extract
        if filename.endswith(".zip") and password:
            try:
                extract_zip_with_password(
                    zip_path=filepath,
                    extract_to=save_dir,
                    password=password
                )
            except BadZipFile:
                print("Gagal membuka ZIP: file rusak atau password salah.")

    if count == 0:
        print("Tidak ada attachment ditemukan.")
    else:
        print(f"Total attachment disimpan: {count}")

def extract_date_from_subject(subject: str) -> Optional[str]:
    """
    Ekstrak tanggal dari subject email dan konversi ke format YYYY-MM-DD.

    Contoh subject:
        - "DATA PICKUP JNE 30/06/25" → "2025-06-30"
        - "OS CARD JNE 01/07/2025"   → "2025-07-01"

    Returns:
        str (format YYYY-MM-DD) jika berhasil, None jika gagal.
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

# mbox_path = "D:/email/Email TB/outlook.office365.com/Inbox"
# buka_thunderbird()
# refresh_inbox(mbox_path)

# print("Membaca file:", mbox_path)

# # Temukan email terbaru yang sesuai dengan kriteria
# latest = find_latest_matching_email(
#     file_path=mbox_path,
#     from_email="jne-dashboard@jne.co.id",
#     subject_prefix="SHIPMENT CCC",
#     max_emails=100
# )

# # Proses hasil
# if latest:
#     body = get_email_body(latest)
    
#     info = extract_login_info_html(body)
#     if info:
#         print("Informasi berhasil ditemukan:")
#         print("Link:", info["Link"])
#         print("Username:", info["Username"])
#         print("Password:", info["Password"])

#         # Simpan ke CSV
#         filename = "login_info.csv"
#         with open(filename, mode="w", newline="", encoding="utf-8") as f:
#             writer = csv.DictWriter(f, fieldnames=["Link", "Username", "Password"])
#             writer.writeheader()
#             writer.writerow(info)

#         print(f"Data berhasil disimpan ke: {filename}")
#     else:
#         print("Tidak ditemukan pola Link, Username, dan Password dalam body email.")