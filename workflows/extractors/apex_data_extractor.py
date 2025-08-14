import os
import time
import logging
from dataclasses import dataclass
from typing import Iterable, List, Optional, Sequence, Tuple

from dotenv import load_dotenv
from datetime import datetime, timedelta
from urllib.parse import unquote
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
)
from pathlib import Path
from tasks.ryan_tasks import open_awb_tasks, rt_awb_tasks

# =============================================================
# Env & Global Config
# =============================================================
project_root = Path(__file__).resolve().parents[2]  # Folder ReportApp
env_path = project_root / ".env"

if env_path.exists():
    load_dotenv(dotenv_path=env_path, override=True)
else:
    raise FileNotFoundError(f"❌ File .env tidak ditemukan di {env_path}")

USERNAME = os.getenv("USERNAME")
PASSWORD = os.getenv("PASSWORD")
APEX_HOSTS = [h.strip() for h in os.getenv("APEX_HOSTS", "").split(",") if h.strip()]

# =============================================================
# Logging Setup
# =============================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("apex_refactor")

# =============================================================
# Dataclasses & Constants
# =============================================================
@dataclass
class RetryPolicy:
    max_attempts: int = 3
    delay_seconds: int = 10


API_ERR_KEYWORDS: Tuple[str, ...] = ("500", "502", "API")
DATE_FMT_VIEW = "%d-%b-%Y"
DATE_FMT_FILE = "%y%m%d"

# =============================================================
# WebDriver Helpers
# =============================================================

def start_driver(download_dir: str) -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_argument("--log-level=3")
    options.add_argument("--headless")  # Jalankan di background
    return webdriver.Chrome(options=options)


def wait_for(driver, condition, timeout: int = 30):
    return WebDriverWait(driver, timeout).until(condition)


def ensure_history_ready(driver, timeout: int = 60):
    """Pastikan halaman History sudah siap dipakai."""
    try:
        # Jika search field sudah ada, selesai
        wait_for(driver, EC.presence_of_element_located((By.ID, "Status History SLA_search_field")), timeout)
        wait_for(driver, EC.element_to_be_clickable((By.ID, "Status History SLA_search_button")), timeout)
        return
    except TimeoutException:
        pass

    # Belum siap → klik tombol History lalu tunggu lagi
    logger.info("Membuka halaman History...")
    click_history_button(driver)
    wait_for(driver, EC.presence_of_element_located((By.ID, "Status History SLA_search_field")), timeout)
    wait_for(driver, EC.element_to_be_clickable((By.ID, "Status History SLA_search_button")), timeout)


# =============================================================
# APEX Actions (Login, Upload, History, Request, Download)
# =============================================================

def login_to_apex(driver, url: str, username: str, password: str):
    driver.get(url)
    time.sleep(1.5)
    logger.info("Login ke sistem...")
    driver.find_element(By.NAME, "p_t01").clear()
    driver.find_element(By.NAME, "p_t01").send_keys(username)
    driver.find_element(By.NAME, "p_t02").send_keys(password + Keys.RETURN)
    time.sleep(1.5)


def upload_file(driver, file_path: str, file_name: str, timeout: int = 5 * 60):
    full_file_path = Path(file_path).joinpath(file_name).resolve()
    logger.info(f"Mengunggah file: {file_name}")

    if not full_file_path.is_file():
        raise FileNotFoundError(f"File tidak ditemukan: {full_file_path}")

    try:
        wait = WebDriverWait(driver, timeout)
        upload_input = wait.until(EC.presence_of_element_located((By.ID, "P45_BLOB_CONTENT")))
        upload_input.send_keys(str(full_file_path))

        submit_button = wait.until(EC.element_to_be_clickable((By.ID, "B49858870221701936")))
        submit_button.click()

        logger.info("File dikirim. Tunggu respon server...")
        time.sleep(3)
        driver.refresh()
    except Exception as e:
        logger.error(f"Gagal upload {file_name}: {e}")
        raise


def click_report_status_awb(driver, timeout: int = 30):
    try:
        link = wait_for(driver, EC.element_to_be_clickable((By.LINK_TEXT, "Report Status AWB")), timeout)
        link.click()
    except Exception as e:
        logger.error(f"Gagal klik tombol 'Report Status AWB': {e}")
        raise


def click_history_button(driver, retries: int = 3):
    for _ in range(retries):
        try:
            btn = wait_for(driver, EC.element_to_be_clickable((By.ID, "B49871477158701971")), 120)
            btn.click()
            return
        except StaleElementReferenceException:
            driver.refresh()
            time.sleep(2)
        except Exception:
            time.sleep(1)
    raise RuntimeError("Gagal klik tombol History setelah beberapa kali percobaan.")


def isi_tanggal_awal_akhir(driver, tanggal_awal: str, tanggal_akhir: str):
    for field_id, value in [("P45_DATE1", tanggal_awal), ("P45_DATE2", tanggal_akhir)]:
        elem = driver.find_element(By.ID, field_id)
        elem.clear()
        elem.send_keys(value)


def fill_form_request(driver, customer_id: str, tanggal_awal: str, tanggal_akhir: str, retries: int = 5):
    for _ in range(retries):
        try:
            try:
                input_cust = driver.find_element(By.ID, "P45_CUST")
                if customer_id.startswith("8"):
                    driver.find_element(By.ID, "P45_NA").send_keys(Keys.ARROW_UP)
                    time.sleep(1)
            except NoSuchElementException:
                driver.find_element(By.ID, "P45_PARAM_TYPE").send_keys(Keys.ARROW_DOWN)
                time.sleep(1)
                if customer_id.startswith("8"):
                    driver.find_element(By.ID, "P45_NA").send_keys(Keys.ARROW_UP)
                    time.sleep(1)
                driver.find_element(By.ID, "P45_PARAM_TYPE").send_keys(Keys.ARROW_DOWN)
                driver.refresh()
                time.sleep(1.5)
                input_cust = driver.find_element(By.ID, "P45_CUST")

            input_cust.clear()
            input_cust.send_keys(customer_id)
            time.sleep(0.5)
            input_cust.send_keys(Keys.ARROW_DOWN)
            input_cust.send_keys(Keys.ENTER)

            isi_tanggal_awal_akhir(driver, tanggal_awal, tanggal_akhir)
            driver.find_element(By.ID, "B49858870221701936").click()
            logger.info(f"Request berhasil untuk {customer_id} ({tanggal_awal} - {tanggal_akhir})")
            return
        except StaleElementReferenceException:
            time.sleep(1)
            driver.refresh()
        except Exception:
            time.sleep(1)
            driver.refresh()
    raise RuntimeError(
        f"Gagal mengisi form request data untuk {customer_id} ({tanggal_awal} - {tanggal_akhir}) setelah {retries}x"
    )


def click_download_link(
    driver,
    *,
    file_prefix: Optional[str] = None,
    customer_id: Optional[str] = None,
    start_fmt: Optional[str] = None,
    end_fmt: Optional[str] = None,
    poll_seconds: int = 10,
    ready_timeout: int = 60,
):
    logger.info("Menunggu link download siap (auto-refresh setiap 10 detik)...")

    if file_prefix:
        search_key = file_prefix
    elif customer_id and start_fmt and end_fmt:
        search_key = f"_{customer_id}_{start_fmt}_{end_fmt}_"
    else:
        raise ValueError("Harus isi file_prefix ATAU customer_id + tanggal")

    # Pastikan halaman History siap
    ensure_history_ready(driver, timeout=ready_timeout)

    while True:
        try:
            search_field = driver.find_element(By.ID, "Status History SLA_search_field")
            search_field.clear()
            search_field.send_keys(search_key)
            time.sleep(1)
            driver.find_element(By.ID, "Status History SLA_search_button").click()
            time.sleep(1.5)

            rows = driver.find_elements(By.CSS_SELECTOR, "table.a-IRR-table tbody tr")
            for row in rows:
                try:
                    filename_text = row.find_element(By.XPATH, './td[@headers="FILENAME"]').text.strip()
                    error_text = row.find_element(By.XPATH, './td[@headers="ERROR"]').text.strip()

                    matched = (file_prefix and filename_text.startswith(file_prefix)) or (
                        customer_id and f"_{customer_id}_{start_fmt}_{end_fmt}_" in filename_text
                    )
                    if not matched:
                        continue

                    if error_text and error_text != "-":
                        if any(x in error_text for x in ["500", "502", "try again"]):
                            raise RuntimeError(f"Proses dihentikan karena error kritikal: {error_text}")
                        if "No Data Found" in error_text:
                            logger.info(f"Tidak ada data untuk {customer_id}, lanjut akun berikutnya.")
                            return None

                    link = row.find_element(By.XPATH, './td[@headers="DOWNLOAD"]/a')
                    href = link.get_attribute("href")
                    actual_filename = unquote(href.split("p_fn=")[-1]) if "p_fn=" in href else filename_text
                    link.click()
                    logger.info(f"Tautan download diklik: {actual_filename}")
                    return actual_filename

                except (NoSuchElementException, StaleElementReferenceException):
                    continue

            time.sleep(poll_seconds)
            driver.refresh()
            ensure_history_ready(driver, timeout=ready_timeout)
        except NoSuchElementException:
            # Jika field tidak ada (halaman berubah), buka History lagi dan lanjut polling
            ensure_history_ready(driver, timeout=ready_timeout)
            continue


def wait_for_download(
    download_dir: str,
    base_file_name: str,
    *,
    timeout: int = 90,
    target_filename: Optional[str] = None,
    contains_text: Optional[str] = None,
) -> str:
    logger.info("Menunggu file selesai diunduh...")
    polling_interval = 2
    elapsed = 0

    while elapsed < timeout:
        for fname in os.listdir(download_dir):
            full_path = os.path.join(download_dir, fname)
            if fname.endswith(".crdownload"):
                continue

            # 1) Cocok persis
            if target_filename and fname == target_filename:
                if os.path.isfile(full_path) and os.path.getsize(full_path) > 0:
                    logger.info(f"File berhasil diunduh: {full_path} ({os.path.getsize(full_path)} bytes)")
                    return full_path
            # 2) Mengandung text (customer_id)
            elif contains_text and contains_text in fname and fname.endswith(".csv"):
                if os.path.isfile(full_path) and os.path.getsize(full_path) > 0:
                    logger.info(f"File berhasil diunduh (by contains): {full_path} ({os.path.getsize(full_path)} bytes)")
                    return full_path
            # 3) Fallback awalan base_file_name
            elif not target_filename and not contains_text:
                if fname.startswith(base_file_name) and fname.endswith(".csv") and os.path.isfile(full_path):
                    if os.path.getsize(full_path) > 0:
                        logger.info(f"File berhasil diunduh: {full_path} ({os.path.getsize(full_path)} bytes)")
                        return full_path

        time.sleep(polling_interval)
        elapsed += polling_interval

    logger.warning("File tidak ditemukan atau belum selesai dalam batas waktu.")
    return ""


# =============================================================
# Utilities (Tanggal, Ranges, Tracker integration)
# =============================================================

def generate_date_ranges(tgl_awal_str: str, tgl_akhir_str: str, fmt: str = DATE_FMT_VIEW, max_range: int = 5) -> List[Tuple[str, str]]:
    tgl_awal = datetime.strptime(tgl_awal_str, fmt)
    tgl_akhir = datetime.strptime(tgl_akhir_str, fmt)
    ranges: List[Tuple[str, str]] = []

    while tgl_awal <= tgl_akhir:
        tgl_end = min(tgl_awal + timedelta(days=max_range - 1), tgl_akhir)
        ranges.append((tgl_awal.strftime(fmt), tgl_end.strftime(fmt)))
        tgl_awal = tgl_end + timedelta(days=1)

    return ranges


def get_tanggal_awal(file_referensi: str, fmt: str = DATE_FMT_VIEW) -> str:
    fallback = datetime.today() - timedelta(days=1)

    if os.path.exists(file_referensi):
        mtime = datetime.fromtimestamp(os.path.getmtime(file_referensi))
        return (fallback if mtime.date() >= fallback.date() else mtime).strftime(fmt)

    return fallback.strftime(fmt)


def all_customers_done(task_items: Sequence, tracker) -> bool:
    for item in task_items:
        if isinstance(item, dict):
            if "nama_customer" in item:
                task_name = item["nama_customer"]
            elif "desc" in item:
                task_name = item["desc"]
            else:
                task_name = str(item)
        else:
            task_name = str(item)

        row = next((r for r in tracker.rows if str(r["task"]) == task_name), None)
        if not row:
            print(f"[DEBUG] Task {task_name} tidak ditemukan di tracker.")
            return False

        download_flag = row.get("download")
        path_flag = str(row.get("path", "")).strip().lower()

        # Skip dianggap selesai
        if str(download_flag).strip().lower() == "skip":
            continue

        if not (download_flag is True or path_flag == "no_data"):
            print(f"[DEBUG] Task {task_name} belum selesai. download={download_flag}, path={path_flag}")
            return False
    return True

# =============================================================
# High-level Orchestration
# =============================================================

def _upload_open_new_rt_if_needed(driver, apex_url_upload: str, tracker, open_awb_tasks, new_awb_tasks, rt_awb_tasks,
                                  file_path: str, file_name: str, download_dir: str, base_file_name: str) -> None:
    # Cek apakah SEMUA task open_awb/new_awb/rt_awb sudah pernah download
    already = True
    for task in open_awb_tasks + new_awb_tasks + rt_awb_tasks:
        tracker_row = next((row for row in tracker.rows if str(row["task"]) == str(task["desc"])), None)
        if not tracker_row or not tracker_row.get("download"):
            already = False
            break

    if already:
        logger.info("Semua file open_awb/new_awb/rt_awb sudah pernah didownload, skip upload & download.")
        return

    # Lakukan upload + unduh hasil open/new/rt
    login_to_apex(driver, apex_url_upload, USERNAME, PASSWORD)
    upload_file(driver, file_path, file_name)

    for task in open_awb_tasks + new_awb_tasks + rt_awb_tasks:
        tracker.set_request(task["desc"], True)

    click_history_button(driver)
    ensure_history_ready(driver)

    downloaded_filename = click_download_link(driver, file_prefix=file_name)
    downloaded_open_file = wait_for_download(download_dir, base_file_name, target_filename=downloaded_filename)

    if downloaded_open_file:
        for task in open_awb_tasks + new_awb_tasks + rt_awb_tasks:
            tracker.set_download(task["desc"], True)
            tracker.set_path(task["desc"], downloaded_open_file)


def _collect_customer_requests(driver, apex_url_request: str, list_customer_ids, tracker,
                               tanggal_awal: str, tanggal_akhir: str) -> List[Tuple[str, str, str, str]]:
    login_to_apex(driver, apex_url_request, USERNAME, PASSWORD)

    request_info: List[Tuple[str, str, str, str]] = []
    for cust in list_customer_ids:
        tracker_row = next((row for row in tracker.rows if str(row["task"]) == str(cust["nama_customer"])), None)
        if tracker_row and (tracker_row.get("download") or tracker_row.get("path") == "no_data"):
            continue

        customer_id = cust["customer_id"]
        include_today = cust.get("include_today", False)
        tanggal_akhir_eff = tanggal_akhir if not include_today else datetime.today().strftime(DATE_FMT_VIEW)

        for start, end in generate_date_ranges(tanggal_awal, tanggal_akhir_eff):
            try:
                fill_form_request(driver, customer_id, start, end)
                tracker.set_request(cust["nama_customer"], True)
                request_info.append(
                    (
                        customer_id,
                        datetime.strptime(start, DATE_FMT_VIEW).strftime(DATE_FMT_FILE),
                        datetime.strptime(end, DATE_FMT_VIEW).strftime(DATE_FMT_FILE),
                        cust["nama_customer"],
                    )
                )
            except Exception as e:
                logger.error(f"Request gagal {customer_id} ({start}-{end}): {e}")

    return request_info


def _download_requested_files(driver, request_info: List[Tuple[str, str, str, str]], download_dir: str, base_file_name: str,
                              tracker, retry: RetryPolicy) -> None:
    click_history_button(driver)
    ensure_history_ready(driver)

    for customer_id, start_fmt, end_fmt, nama_customer in request_info:
        row = next((r for r in tracker.rows if str(r["task"]) == nama_customer), None)
        if row and row.get("download"):
            continue

        start_date = datetime.strptime(start_fmt, DATE_FMT_FILE).strftime(DATE_FMT_VIEW)
        end_date = datetime.strptime(end_fmt, DATE_FMT_FILE).strftime(DATE_FMT_VIEW)

        for attempt in range(1, retry.max_attempts + 1):
            try:
                filename = click_download_link(
                    driver,
                    customer_id=customer_id,
                    start_fmt=start_fmt,
                    end_fmt=end_fmt,
                )

                # None → No Data
                if filename is None:
                    tracker.set_download(nama_customer, "Skip")
                    tracker.set_path(nama_customer, "no_data")
                    break

                downloaded = wait_for_download(
                    download_dir,
                    base_file_name,
                    target_filename=filename,
                    contains_text=customer_id,
                )

                if downloaded:
                    tracker.set_download(nama_customer, True)
                    tracker.set_path(nama_customer, downloaded)

                    # ✅ Jika nama file sama dengan base_file_name → set semua task open_awb_tasks + rt_awb_tasks
                    if base_file_name.lower() in os.path.basename(downloaded).lower():
                        for t in list(open_awb_tasks) + list(rt_awb_tasks):
                            task_name = t["nama_customer"] if isinstance(t, dict) else str(t)
                            tracker.set_download(task_name, True)
                            tracker.set_path(task_name, downloaded)

                    break

                raise RuntimeError("File belum berhasil di-download.")

            except Exception as e:
                err = str(e)
                # API error → re-request lalu retry
                if any(k in err for k in API_ERR_KEYWORDS):
                    logger.warning(
                        f"API error (500/502/API) attempt {attempt}/{retry.max_attempts} → {customer_id}"
                    )
                    if attempt < retry.max_attempts:
                        try:
                            click_report_status_awb(driver)
                            fill_form_request(
                                driver,
                                customer_id=customer_id,
                                tanggal_awal=start_date,
                                tanggal_akhir=end_date,
                            )
                            logger.info("↻ Request ulang berhasil, akan coba download lagi...")
                        except Exception as req_e:
                            logger.error(f"Gagal request ulang: {req_e}")
                            break
                        time.sleep(retry.delay_seconds)
                        continue
                # Error lain → keluar
                logger.error(f"Download gagal untuk {customer_id}: {err}")
                break

        tracker.summary(print_summary=False)


# =============================================================
# Public API (entry point)
# =============================================================

def process_apex_upload_and_request(
    file_name,
    base_file_name,
    file_path,
    download_dir,
    list_customer_ids,
    tracker,
    open_awb_tasks,
    new_awb_tasks,
    rt_awb_tasks,
):
    """
    Orkestrasi penuh:
    1) Untuk setiap host APEX:
       - (Opsional) upload open/new/rt bila belum pernah didownload semuanya
       - Kirim request data per customer (auto split tanggal)
       - Masuk History, unduh semua hasil
    2) Hentikan lebih awal jika semua task selesai.
    """
    tanggal_awal = get_tanggal_awal(r"C:\\Users\\DELL\\Desktop\\ReportApp\\data\\archive\\Open AWB Danamon.csv")
    yesterday_str = (datetime.today() - timedelta(days=1)).strftime(DATE_FMT_VIEW)

    retry_policy = RetryPolicy(max_attempts=3, delay_seconds=10)
    success = False

    combined_tasks = list(open_awb_tasks) + list(new_awb_tasks) + list(rt_awb_tasks) + list(list_customer_ids)

    for host in APEX_HOSTS:
        # Early exit sebelum mulai host
        if all_customers_done(combined_tasks, tracker):
            logger.info(f"Semua task sudah selesai sebelum masuk host: {host}")
            success = True
            break

        logger.info(f"Mencoba koneksi ke APEX host: {host}")
        apex_url_upload = f"http://{host}:7777/apex/f?p=105:45:::NO::P45_PARAM_TYPE:A"
        apex_url_request = f"http://{host}:7777/apex/f?p=105:45:::NO::P45_PARAM_TYPE:P"

        driver = start_driver(download_dir)
        try:
            # 1) Upload open/new/rt bila perlu
            _upload_open_new_rt_if_needed(
                driver,
                apex_url_upload,
                tracker,
                open_awb_tasks,
                new_awb_tasks,
                rt_awb_tasks,
                file_path,
                file_name,
                download_dir,
                base_file_name,
            )

            # 2) Kumpulkan request untuk semua customer
            request_info = _collect_customer_requests(
                driver,
                apex_url_request,
                list_customer_ids,
                tracker,
                tanggal_awal,
                yesterday_str,
            )

            # 3) Download semua hasil request dari History
            _download_requested_files(
                driver,
                request_info,
                download_dir,
                base_file_name,
                tracker,
                retry_policy,
            )

            # 4) Cek lagi setelah semua proses di host
            if all_customers_done(combined_tasks, tracker):
                logger.info(f"Semua task sudah selesai di host: {host}")
                success = True
                break

        except Exception as e:
            logger.error(f"Error di host {host}: {e}")
            # lanjut ke host berikutnya
        finally:
            driver.quit()
            logger.info("Browser ditutup.")

    return success
