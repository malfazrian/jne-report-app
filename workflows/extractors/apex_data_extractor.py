import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException

def start_driver(download_dir: str):
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=options)

def login_to_apex(driver, url: str, username: str, password: str):
    driver.get(url)
    time.sleep(2)
    print("Login ke sistem...")
    driver.find_element(By.NAME, "p_t01").send_keys(username)
    driver.find_element(By.NAME, "p_t02").send_keys(password + Keys.RETURN)

def upload_file(driver, file_path: str, file_name: str):
    time.sleep(2)
    print("Mengunggah file:", file_name)
    driver.find_element(By.ID, "P45_BLOB_CONTENT").send_keys(file_path)
    driver.find_element(By.ID, "B49858870221701936").click()
    time.sleep(5)
    driver.refresh()

def click_download_link(driver, file_name: str):
    print("Menunggu link download siap (auto-refresh setiap 10 detik)...")
    while True:
        rows = driver.find_elements(By.CSS_SELECTOR, "tr.highlight-row")
        for row in rows:
            if file_name in row.text:
                try:
                    error_cell = row.find_element(By.XPATH, './td[@headers="ERROR"]')
                    error_text = error_cell.text.strip()
                    if error_text and error_text.lower() != "-":
                        print(f"Gagal proses file: {error_text}")
                        raise SystemExit()
                except NoSuchElementException:
                    pass

                try:
                    download_cell = row.find_element(By.XPATH, './td[@headers="DOWNLOAD"]')
                    link = download_cell.find_element(By.TAG_NAME, "a")
                    link.click()
                    print(f"Tautan download diklik: {file_name}")
                    return
                except NoSuchElementException:
                    pass
        time.sleep(10)
        driver.refresh()

def wait_for_download(download_dir: str, base_file_name: str, timeout: int = 90) -> str:
    print("Menunggu file selesai diunduh...")
    polling_interval = 2
    elapsed = 0

    while elapsed < timeout:
        for fname in os.listdir(download_dir):
            if fname.startswith(base_file_name) and fname.endswith(".csv"):
                if fname.endswith(".crdownload"):
                    continue  # masih proses download

                full_path = os.path.join(download_dir, fname)
                if os.path.isfile(full_path) and os.path.getsize(full_path) > 0:
                    crdownload_path = full_path + ".crdownload"
                    if not os.path.exists(crdownload_path):  # pastikan download benar-benar selesai
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

def isi_form_request(driver, customer_id: str, tanggal_awal: str, tanggal_akhir: str, retries: int = 2):
    for attempt in range(retries):
        try:
            try:
                input_cust = driver.find_element(By.ID, "P45_CUST")
            except NoSuchElementException:
                print("⚠️ Form belum muncul, mencoba shortcut + refresh...")
                driver.find_element(By.ID, "P45_PARAM_TYPE").send_keys(Keys.ARROW_DOWN)
                time.sleep(2)
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
    raise Exception("Gagal mengisi form request data setelah beberapa percobaan")

def generate_date_ranges(tgl_awal_str: str, tgl_akhir_str: str, format: str = "%d-%b-%Y", max_range: int = 5):
    tgl_awal = datetime.datetime.strptime(tgl_awal_str, format)
    tgl_akhir = datetime.datetime.strptime(tgl_akhir_str, format)
    ranges = []

    while tgl_awal <= tgl_akhir:
        tgl_end = min(tgl_awal + datetime.timedelta(days=max_range - 1), tgl_akhir)
        ranges.append((
            tgl_awal.strftime(format),
            tgl_end.strftime(format)
        ))
        tgl_awal = tgl_end + datetime.timedelta(days=1)

    return ranges