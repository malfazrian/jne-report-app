import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# === SETTING ===
WHATSAPP_PROFILE = {
    'user_data_dir': r'C:\Users\DELL\AppData\Local\Selenium',
    'profile_directory': 'Profile 3'
}

SEND_DELAY = 2  # jeda antar file (detik)
TIMEOUT = 60    # waktu maksimal tunggu (detik)

# FILES_TO_SEND = [
#     {
#         'contact': 'Abidin',
#         'files': [
#             r'D:\RYAN\3. Reports\Generali\Generali Mei 2025.xlsx',
#             r'D:\RYAN\3. Reports\Generali\Generali Juni 2025.xlsx',
#             r'D:\RYAN\3. Reports\Generali\Generali Juli 2025.xlsx'
#         ]
#     }
# ]

# === FUNGSI: Inisialisasi driver Chrome ===
def init_driver():
    opts = Options()
    opts.add_argument(f"--user-data-dir={WHATSAPP_PROFILE['user_data_dir']}")
    opts.add_argument(f"--profile-directory={WHATSAPP_PROFILE['profile_directory']}")
    opts.add_argument('--disable-extensions')
    opts.add_argument('--disable-blink-features=AutomationControlled')
    opts.add_argument('--window-size=1920,1080')
    driver = webdriver.Chrome(options=opts)
    driver.get('https://web.whatsapp.com/')
    return driver, WebDriverWait(driver, TIMEOUT)

# === FUNGSI: Kirim 1 file ke 1 kontak ===
def send_file_to_contact(wait, contact_name, file_path):
    if not os.path.exists(file_path):
        print(f"File tidak ditemukan: {file_path}")
        return

    print(f"Mengirim ke: {contact_name} → {os.path.basename(file_path)}")

    # Cari dan klik kontak
    search_box = wait.until(EC.presence_of_element_located((
        By.XPATH, "//div[@role='textbox' and @aria-placeholder='Search or start a new chat']"
    )))
    search_box.click()
    search_box.clear()
    search_box.send_keys(contact_name)
    time.sleep(2)

    chat = wait.until(EC.element_to_be_clickable((
        By.XPATH, f"//span[@title='{contact_name}']"
    )))
    chat.click()

    # Klik tombol attach dan upload file
    wait.until(EC.element_to_be_clickable((
        By.XPATH, "//button[@title='Attach']"
    ))).click()

    upload_input = wait.until(EC.presence_of_element_located((
        By.XPATH, "//input[@accept='*']"
    )))
    upload_input.send_keys(file_path)

    wait.until(EC.element_to_be_clickable((
        By.XPATH, "//div[@role='button' and @aria-label='Send']"
    ))).click()

    print("Menunggu status terkirim...")
    wait.until(EC.presence_of_element_located((
        By.XPATH, '//span[@aria-label=" Read " or contains(@aria-label, "Delivered")]'
    )))
    print(f"Terkirim & status terdeteksi: {os.path.basename(file_path)}")

# === FUNGSI: Loop pengiriman untuk semua kontak & file ===
def send_all_files(file_list):
    driver, wait = init_driver()
    print("Menunggu WhatsApp Web terbuka...\n")

    for entry in file_list:
        contact = entry['contact']
        for file_path in entry['files']:
            try:
                send_file_to_contact(wait, contact, file_path)
                time.sleep(SEND_DELAY)
            except Exception as e:
                print(f"❌ Gagal kirim ke {contact}: {e}")
    time.sleep(5)
    driver.quit()
    print("Semua pengiriman selesai.")