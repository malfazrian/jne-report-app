import os
import time
from datetime import datetime, timedelta
from urllib.parse import unquote
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from pathlib import Path

apex_hosts = ["10.18.2.35", "10.18.2.16", "10.18.2.12", "10.18.2.11"]
username = "ccc.support4@jne.co.id"
password = "123"
success = False

def start_driver(download_dir: str):
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=options)

def login_to_apex(driver, url, username, password):
    driver.get(url)
    time.sleep(2)
    print("Login ke sistem...")
    driver.find_element(By.NAME, "p_t01").clear()
    driver.find_element(By.NAME, "p_t01").send_keys(username)
    driver.find_element(By.NAME, "p_t02").send_keys(password + Keys.RETURN)


def upload_file(driver, file_path: str, file_name: str):
    full_file_path = Path(file_path).joinpath(file_name).resolve()
    print(f"Mengunggah file: {file_name}")

    if not full_file_path.is_file():
        raise FileNotFoundError(f"File tidak ditemukan: {full_file_path}")

    upload_input = driver.find_element(By.ID, "P45_BLOB_CONTENT")
    upload_input.send_keys(str(full_file_path))

    driver.find_element(By.ID, "B49858870221701936").click()
    time.sleep(5)
    driver.refresh()

def upload_all_csv_in_folder(driver, folder_path: str):
    # Ambil semua file .csv di folder
    csv_files = [f for f in os.listdir(folder_path) if f.endswith(".csv")]
    
    if not csv_files:
        print("Tidak ada file .csv ditemukan di folder:", folder_path)
        return

    for csv_file in csv_files:
        full_path = os.path.join(folder_path, csv_file)
        try:
            upload_file(driver, full_path, csv_file)
        except Exception as e:
            print(f"Gagal upload {csv_file}: {e}")

def click_download_link(driver, file_prefix: str = None, customer_id: str = None, start_fmt: str = None, end_fmt: str = None):
    print("Menunggu link download siap (auto-refresh setiap 10 detik)...")

    if file_prefix:
        search_key = file_prefix
    elif customer_id and start_fmt and end_fmt:
        partial_key = f"_{customer_id}_{start_fmt}_{end_fmt}_"
        search_key = partial_key
    else:
        raise ValueError("Harus isi file_prefix ATAU customer_id + tanggal")

    partial_key = f"_{customer_id}_{start_fmt}_{end_fmt}_" if customer_id else None

    while True:
        # Clear search field
        search_field = driver.find_element(By.ID, "Status History SLA_search_field")
        search_field.clear()
        search_field.send_keys(search_key)
        time.sleep(2)
        driver.find_element(By.ID, "Status History SLA_search_button").click()
        time.sleep(2)

        rows = driver.find_elements(By.CSS_SELECTOR, "table.a-IRR-table tbody tr")
        for row in rows:
            try:
                filename_cell = row.find_element(By.XPATH, './td[@headers="FILENAME"]')
                filename_text = filename_cell.text.strip()

                if file_prefix and filename_text.startswith(file_prefix):
                    matched = True
                elif partial_key and partial_key in filename_text:
                    matched = True
                else:
                    matched = False

                if not matched:
                    continue

                # Cek error
                error_cell = row.find_element(By.XPATH, './td[@headers="ERROR"]')
                error_text = error_cell.text.strip()
                if error_text and error_text != "-":
                    print(f"Gagal proses file: {error_text}")
                    raise Exception(f"File error: {error_text}")

                # Klik link download
                download_cell = row.find_element(By.XPATH, './td[@headers="DOWNLOAD"]')
                link = download_cell.find_element(By.TAG_NAME, "a")
                href = link.get_attribute("href")

                # Ekstrak nama file dari href, misal: REPCCC.GET_FILE?p_fn=Update Ryan 040625_6383809.csv
                if "p_fn=" in href:
                    actual_filename = unquote(href.split("p_fn=")[-1])
                else:
                    actual_filename = filename_text  # fallback

                link.click()
                print(f"Tautan download diklik: {actual_filename}")

                return actual_filename

            except NoSuchElementException:
                continue

        time.sleep(10)
        driver.refresh()

def wait_for_download(download_dir: str, base_file_name: str, timeout: int = 90, target_filename=None) -> str:
    print("Menunggu file selesai diunduh...")
    polling_interval = 2
    elapsed = 0

    while elapsed < timeout:
        for fname in os.listdir(download_dir):
            if target_filename:
                if fname != target_filename:
                    continue
            else:
                if not fname.startswith(base_file_name) or not fname.endswith(".csv"):
                    continue
                if fname.endswith(".crdownload"):
                    continue

            full_path = os.path.join(download_dir, fname)
            if os.path.isfile(full_path) and os.path.getsize(full_path) > 0:
                crdownload_path = full_path + ".crdownload"
                if not os.path.exists(crdownload_path):
                    print(f"File berhasil diunduh: {full_path} ({os.path.getsize(full_path)} bytes)")
                    return full_path

        time.sleep(polling_interval)
        elapsed += polling_interval

    print("File tidak ditemukan atau belum selesai dalam batas waktu.")
    return ""

def isi_tanggal_awal_akhir(driver, tanggal_awal: str, tanggal_akhir: str):
    for field_id, value in [("P45_DATE1", tanggal_awal), ("P45_DATE2", tanggal_akhir)]:
        elem = driver.find_element(By.ID, field_id)
        elem.clear()
        elem.send_keys(value)

def fill_form_request(driver, customer_id: str, tanggal_awal: str, tanggal_akhir: str, retries: int = 5):
    for attempt in range(retries):
        try:
            try:
                input_cust = driver.find_element(By.ID, "P45_CUST")
                if customer_id.startswith("8"):
                    driver.find_element(By.ID, "P45_NA").send_keys(Keys.ARROW_UP)
                    time.sleep(2)
            except NoSuchElementException:
                print("Form belum muncul, mencoba shortcut + refresh...")
                driver.find_element(By.ID, "P45_PARAM_TYPE").send_keys(Keys.ARROW_DOWN)
                time.sleep(2)
                if customer_id.startswith("8"):
                    driver.find_element(By.ID, "P45_NA").send_keys(Keys.ARROW_UP)
                    time.sleep(2)
                driver.find_element(By.ID, "P45_PARAM_TYPE").send_keys(Keys.ARROW_DOWN)
                driver.refresh()
                time.sleep(3)
                input_cust = driver.find_element(By.ID, "P45_CUST")

            input_cust.clear()
            input_cust.send_keys(customer_id)
            time.sleep(1)
            input_cust.send_keys(Keys.ARROW_DOWN)
            input_cust.send_keys(Keys.ENTER)

            isi_tanggal_awal_akhir(driver, tanggal_awal, tanggal_akhir)
            driver.find_element(By.ID, "B49858870221701936").click()
            time.sleep(5)
            driver.refresh()
            return
        except Exception as e:
            print(f"Percobaan ke-{attempt+1} gagal isi form: {e}")
            time.sleep(2)
            driver.refresh()
    raise Exception("Gagal mengisi form request data setelah 5 kali percobaan")

def generate_date_ranges(tgl_awal_str: str, tgl_akhir_str: str, format: str = "%d-%b-%Y", max_range: int = 5):
    tgl_awal = datetime.strptime(tgl_awal_str, format)
    tgl_akhir = datetime.strptime(tgl_akhir_str, format)
    ranges = []

    while tgl_awal <= tgl_akhir:
        tgl_end = min(tgl_awal + timedelta(days=max_range - 1), tgl_akhir)
        ranges.append((
            tgl_awal.strftime(format),
            tgl_end.strftime(format)
        ))
        tgl_awal = tgl_end + timedelta(days=1)

    return ranges

def get_tanggal_awal(file_referensi: str, format: str = "%d-%b-%Y"):
    fallback = datetime.today() - timedelta(days=1)
    
    if os.path.exists(file_referensi):
        mtime = datetime.fromtimestamp(os.path.getmtime(file_referensi))

        # Kalau file dimodifikasi hari ini atau lebih baru dari fallback, pakai fallback
        if mtime.date() >= fallback.date():
            return fallback.strftime(format)
        
        return mtime.strftime(format)

    return fallback.strftime(format)

def process_apex_upload_and_request(file_name, base_file_name, file_path, download_dir, list_customer_ids):
    list_customer = [
        {
            "customer_id": cust["customer_id"],
            "include_today": cust.get("include_today", False)
        }
        for cust in list_customer_ids
    ]
    archive_file = r"C:\Users\DELL\Desktop\ReportApp\data\Archive\Open AWB Smartfren.csv"

    tanggal_awal = get_tanggal_awal(archive_file)
    if not tanggal_awal:
        print("Tidak dapat menentukan tanggal awal, menggunakan kemarin.")
        tanggal_awal = (datetime.today() - timedelta(days=1)).strftime("%d-%b-%Y")
    today_str = datetime.today().strftime("%d-%b-%Y")
    yesterday_str = (datetime.today() - timedelta(days=1)).strftime("%d-%b-%Y")

    for host in apex_hosts:
        print(f"Mencoba koneksi ke APEX host: {host}")
        apex_url_upload = f"http://{host}:7777/apex/f?p=105:45:::NO::P45_PARAM_TYPE:A"
        apex_url_request = f"http://{host}:7777/apex/f?p=105:45:::NO::P45_PARAM_TYPE:P"

        driver = start_driver(download_dir)
        request_info = []
        try:
            login_to_apex(driver, apex_url_upload, username, password)
            print(rf"{file_path} with {file_name} akan diunggah ke APEX...")
            upload_file(driver, file_path, file_name)

            # klik tombol History
            tombol_history = driver.find_element(By.ID, "B49871477158701971")
            tombol_history.click()
            time.sleep(2)

            downloaded_filename = click_download_link(driver, file_prefix=file_name)
            downloaded_file = wait_for_download(download_dir, base_file_name, target_filename=downloaded_filename)

            if not downloaded_file:
                print("File upload tidak berhasil, lanjut host berikutnya.")
                driver.quit()
                continue

            print("Upload dan download sukses.")

            login_to_apex(driver, apex_url_upload, username, password)

            for cust in list_customer:
                customer_id = cust["customer_id"]
                include_today = cust["include_today"]

                tanggal_akhir = today_str if include_today else yesterday_str
                date_ranges = generate_date_ranges(tanggal_awal, tanggal_akhir)

                for start, end in date_ranges:
                    try:
                        fill_form_request(driver, customer_id, start, end)
                        request_info.append((
                            customer_id,
                            datetime.strptime(start, "%d-%b-%Y").strftime("%y%m%d"),
                            datetime.strptime(end, "%d-%b-%Y").strftime("%y%m%d"),
                        ))
                    except Exception as e:
                        print(f"Request gagal {customer_id} ({start}-{end}): {e}")

            
            # klik tombol History
            tombol_history = driver.find_element(By.ID, "B49871477158701971")
            tombol_history.click()
            time.sleep(2)

            for customer_id, start_fmt, end_fmt in request_info:
                try:
                    filename = click_download_link(driver, customer_id=customer_id, start_fmt=start_fmt, end_fmt=end_fmt)
                    wait_for_download(download_dir, base_file_name, target_filename=filename)
                    driver.refresh()
                except Exception as e:
                    print(f"Download gagal {customer_id} ({start_fmt}-{end_fmt}): {e}")

            print("Semua proses APEX berhasil.")
            return True

        except Exception as e:
            print(f"Error di host {host}: {e}")
            continue

        finally:
            driver.quit()
            print("Browser ditutup.")

    return False
