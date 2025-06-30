import os
import time
from datetime import datetime, timedelta
from urllib.parse import unquote
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from pathlib import Path
from workflows.tracker import TaskTableTracker

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
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument("--log-level=3")
    options.add_argument("--headless")
    return webdriver.Chrome(options=options)

def login_to_apex(driver, url, username, password):
    driver.get(url)
    time.sleep(2)
    print("Login ke sistem...")
    driver.find_element(By.NAME, "p_t01").clear()
    driver.find_element(By.NAME, "p_t01").send_keys(username)
    driver.find_element(By.NAME, "p_t02").send_keys(password + Keys.RETURN)
    time.sleep(2)

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

def click_download_link(driver, file_prefix: str = None, customer_id: str = None, start_fmt: str = None, end_fmt: str = None):
    print("Menunggu link download siap (auto-refresh setiap 10 detik)...")

    if file_prefix:
        search_key = file_prefix
    elif customer_id and start_fmt and end_fmt:
        search_key = f"_{customer_id}_{start_fmt}_{end_fmt}_"
    else:
        raise ValueError("Harus isi file_prefix ATAU customer_id + tanggal")

    while True:
        search_field = driver.find_element(By.ID, "Status History SLA_search_field")
        search_field.clear()
        search_field.send_keys(search_key)
        time.sleep(2)
        driver.find_element(By.ID, "Status History SLA_search_button").click()
        time.sleep(2)

        rows = driver.find_elements(By.CSS_SELECTOR, "table.a-IRR-table tbody tr")
        for row in rows:
            try:
                filename_text = row.find_element(By.XPATH, './td[@headers="FILENAME"]').text.strip()
                error_text = row.find_element(By.XPATH, './td[@headers="ERROR"]').text.strip()

                matched = (file_prefix and filename_text.startswith(file_prefix)) or \
                          (customer_id and f"_{customer_id}_{start_fmt}_{end_fmt}_" in filename_text)

                if not matched:
                    continue

                if error_text and error_text != "-":
                    if any(x in error_text for x in ["500", "502", "try again"]):
                        raise Exception(f"Proses dihentikan karena error kritikal: {error_text}")
                    if "No Data Found" in error_text:
                        print(f"Tidak ada data untuk {customer_id}, lanjut akun berikutnya.")
                        return None

                link = row.find_element(By.XPATH, './td[@headers="DOWNLOAD"]').find_element(By.TAG_NAME, "a")
                href = link.get_attribute("href")
                actual_filename = unquote(href.split("p_fn=")[-1]) if "p_fn=" in href else filename_text
                link.click()
                print(f"Tautan download diklik: {actual_filename}")
                return actual_filename

            except (NoSuchElementException, StaleElementReferenceException):
                continue

        time.sleep(10)
        driver.refresh()

def wait_for_download(download_dir: str, base_file_name: str, timeout: int = 90, target_filename=None, contains_text=None) -> str:
    print("Menunggu file selesai diunduh...")
    polling_interval = 2
    elapsed = 0

    while elapsed < timeout:
        for fname in os.listdir(download_dir):
            # Jika ada target_filename, tetap cek persis
            if target_filename and fname == target_filename:
                full_path = os.path.join(download_dir, fname)
                if os.path.isfile(full_path) and os.path.getsize(full_path) > 0 and not os.path.exists(full_path + ".crdownload"):
                    print(f"File berhasil diunduh: {full_path} ({os.path.getsize(full_path)} bytes)")
                    return full_path
            # Jika contains_text (misal customer_id), cek apakah ada di nama file
            elif contains_text and contains_text in fname and fname.endswith(".csv") and not fname.endswith(".crdownload"):
                full_path = os.path.join(download_dir, fname)
                if os.path.isfile(full_path) and os.path.getsize(full_path) > 0 and not os.path.exists(full_path + ".crdownload"):
                    print(f"File berhasil diunduh (by contains): {full_path} ({os.path.getsize(full_path)} bytes)")
                    return full_path
            # Fallback: cek nama file yang diawali base_file_name
            elif not target_filename and not contains_text:
                if not fname.startswith(base_file_name) or not fname.endswith(".csv") or fname.endswith(".crdownload"):
                    continue
                full_path = os.path.join(download_dir, fname)
                if os.path.isfile(full_path) and os.path.getsize(full_path) > 0 and not os.path.exists(full_path + ".crdownload"):
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
            print(f"Request berhasil untuk {customer_id} ({tanggal_awal} - {tanggal_akhir})")
            return
        except StaleElementReferenceException:
            print("StaleElementReferenceException terjadi, cari ulang elemen...")
            time.sleep(2)
            driver.refresh()
        except Exception as e:
            time.sleep(2)
            driver.refresh()
    raise Exception(f"Gagal mengisi form request data untuk {customer_id} ({tanggal_awal} - {tanggal_akhir}) setelah {retries} kali percobaan")

def generate_date_ranges(tgl_awal_str: str, tgl_akhir_str: str, format: str = "%d-%b-%Y", max_range: int = 5):
    tgl_awal = datetime.strptime(tgl_awal_str, format)
    tgl_akhir = datetime.strptime(tgl_akhir_str, format)
    ranges = []

    while tgl_awal <= tgl_akhir:
        tgl_end = min(tgl_awal + timedelta(days=max_range - 1), tgl_akhir)
        ranges.append((tgl_awal.strftime(format), tgl_end.strftime(format)))
        tgl_awal = tgl_end + timedelta(days=1)

    return ranges

def get_tanggal_awal(file_referensi: str, format: str = "%d-%b-%Y"):
    fallback = datetime.today() - timedelta(days=1)

    if os.path.exists(file_referensi):
        mtime = datetime.fromtimestamp(os.path.getmtime(file_referensi))
        if mtime.date() >= fallback.date():
            return fallback.strftime(format)
        return mtime.strftime(format)

    return fallback.strftime(format)

def process_apex_upload_and_request(file_name, base_file_name, file_path, download_dir, list_customer_ids, tracker,
    open_awb_tasks, new_awb_tasks, rt_awb_tasks):
    tanggal_awal = get_tanggal_awal(r"C:/Users/DELL/Desktop/ReportApp/data/Archive/Open AWB Smartfren.csv")
    today_str = datetime.today().strftime("%d-%b-%Y")
    yesterday_str = (datetime.today() - timedelta(days=1)).strftime("%d-%b-%Y")

    for host in apex_hosts:
        print(f"Mencoba koneksi ke APEX host: {host}")
        apex_url_upload = f"http://{host}:7777/apex/f?p=105:45:::NO::P45_PARAM_TYPE:A"
        apex_url_request = f"http://{host}:7777/apex/f?p=105:45:::NO::P45_PARAM_TYPE:P"

        driver = start_driver(download_dir)
        request_info = []

        try:
            # Cek apakah SEMUA task open_awb/new_awb/rt_awb sudah pernah download
            all_done = True
            for task in open_awb_tasks + new_awb_tasks + rt_awb_tasks:
                tracker_row = next((row for row in tracker.rows if str(row["task"]) == str(task["desc"])), None)
                if not tracker_row or not tracker_row.get("download"):
                    all_done = False
                    break

            if all_done:
                print("Semua file open_awb/new_awb/rt_awb sudah pernah didownload, skip upload & download.")
            else:
                login_to_apex(driver, apex_url_upload, username, password)
                upload_file(driver, file_path, file_name)
                for task in open_awb_tasks + new_awb_tasks + rt_awb_tasks:
                    tracker.set_request(task["desc"], True)
                driver.find_element(By.ID, "B49871477158701971").click()
                time.sleep(2)

                downloaded_filename = click_download_link(driver, file_prefix=file_name)
                downloaded_open_file = wait_for_download(download_dir, base_file_name, target_filename=downloaded_filename)

                if downloaded_open_file:
                    for task in open_awb_tasks + new_awb_tasks + rt_awb_tasks:
                        tracker.set_download(task["desc"], True)
                        tracker.set_path(task["desc"], downloaded_open_file)

            login_to_apex(driver, apex_url_request, username, password)

            for cust in list_customer_ids:
                # Cek tracker, jika sudah download, skip
                tracker_row = next((row for row in tracker.rows if str(row["task"]) == str(cust["nama_customer"])), None)
                if tracker_row and tracker_row.get("download"):
                    print(f"Sudah berhasil download untuk {cust['nama_customer']}, skip request & download.")
                    continue

                customer_id = cust["customer_id"]
                include_today = cust.get("include_today", False)
                tanggal_akhir = today_str if include_today else yesterday_str
                date_ranges = generate_date_ranges(tanggal_awal, tanggal_akhir)
                for start, end in date_ranges:
                    try:
                        fill_form_request(driver, customer_id, start, end)
                        tracker.set_request(cust["nama_customer"], True)
                        request_info.append((
                            customer_id,
                            datetime.strptime(start, "%d-%b-%Y").strftime("%y%m%d"),
                            datetime.strptime(end, "%d-%b-%Y").strftime("%y%m%d"),
                            cust["nama_customer"]
                        ))
                    except Exception as e:
                        print(f"Request gagal {customer_id} ({start}-{end}): {e}")

            driver.find_element(By.ID, "B49871477158701971").click()
            time.sleep(2)

            for customer_id, start_fmt, end_fmt, nama_customer in request_info:
                # Cek tracker, jika sudah download, skip
                tracker_row = next((row for row in tracker.rows if str(row["task"]) == str(nama_customer)), None)
                if tracker_row and tracker_row.get("download"):
                    print(f"Sudah berhasil download untuk {nama_customer}, skip download.")
                    continue

                try:
                    filename = click_download_link(driver, customer_id=customer_id, start_fmt=start_fmt, end_fmt=end_fmt)
                    if filename:
                        downloaded = wait_for_download(download_dir, base_file_name, target_filename=filename, contains_text=customer_id)
                        if downloaded:
                            tracker.set_download(nama_customer, True)
                            tracker.set_path(nama_customer, downloaded)
                except Exception as e:
                    print(f"Download gagal untuk {customer_id}: {e}")

            return True

        except Exception as e:
            print(f"Error di host {host}: {e}")
            continue

        finally:
            driver.quit()
            print("Browser ditutup.")

    return False
