# main.py

from workflows.backup_open_awb import backup_open_awb_files
from workflows.extractors.open_awb_extractor import extract_open_awb
from workflows.extractors.new_awb_extractor import extract_new_awb_smartfren
from workflows.extractors.awb_rt_extractor import extract_rt_awb
from core.file_ops import merge_csv_files
from workflows.extractors.apex_data_extractor import (
    start_driver, login_to_apex, upload_file,
    click_download_link, wait_for_download
)
from datetime import datetime
import os

def main():
    print("Memulai ReportApp..")
    
    # 1 Backup file open AWB ke Archive
    print("Menjalankan proses backup file open AWB...")
    backup_open_awb_files()
    
    # 2 Ekstrak open AWB Danamon
    extract_open_awb(
    input_path="D:/RYAN/3. Reports/Danamon/Raw Report.csv",
    output_path="c:/Users/DELL/Desktop/ReportApp/data/Open AWB Danamon.csv",
    file_type="csv",
    status_column="STATUS_POD",
    awb_column="AWB",
    exclude_statuses=["SUCCESS", "RETURN SHIPPER", "DELIVERED"]
    )


    # 3 Ekstrak RT AWB Danamon
    print("Mengekstrak RT AWB dari Danamon...")
    extract_rt_awb(
        input_path="D:/RYAN/3. Reports/Danamon/Raw Report.csv",
        output_path="c:/Users/DELL/Desktop/ReportApp/data/RT AWB Danamon.csv"
    )

    # 4 Ekstrak open AWB Generali
    print(" Mengekstrak Open AWB dari Generali...")
    extract_open_awb(
    input_path="D:/RYAN/3. Reports/Generali",
    output_path="c:/Users/DELL/Desktop/ReportApp/data/Open AWB Generali.csv",
    file_type="excel",
    status_column="DELIVERY STATUS",
    awb_column="AWB",
    exclude_statuses=["SUCCESS", "RETURN SHIPPER"],
    header_row=2
    )


    print("Mengekstrak Open AWB Kalog...")
    extract_open_awb(
    input_path="D:/RYAN/3. Reports/Kalog",
    output_path="c:/Users/DELL/Desktop/ReportApp/data/Open AWB Kalog.csv",
    file_type="excel",
    status_column="STATUS_POD",
    awb_column="AWB",
    exclude_statuses=["SUCCESS", "RETURN SHIPPER", "DELIVERED"]
    )

    print("Mengekstrak Open AWB LSIN...")
    extract_open_awb(
    input_path=r"D:\RYAN\3. Reports\LSIN",
    output_path=r"c:/Users/DELL/Desktop/ReportApp/data/Open AWB LSIN.csv",
    file_type="excel",
    status_column="STATUS_POD",
    awb_column="AWB",
    exclude_statuses=["SUCCESS", "RETURN SHIPPER", "DELIVERED"],
    )

    print("Mengekstrak Open AWB Smartfren...")
    extract_open_awb(
    input_path=r"D:\RYAN\3. Reports\Smartfren",
    output_path=r"c:/Users/DELL/Desktop/ReportApp/data/Open AWB Smartfren.csv",
    file_type="excel",
    status_column="STATUS_POD",
    awb_column="AWB",
    exclude_statuses=["SUCCESS", "RETURN SHIPPER", "DELIVERED"],
    quote_awb=True,
    process_all_sheets=True
    )

    # print("Mengekstrak New AWB Smartfren...")
    # extract_new_awb_smartfren(
    # base_folder=r"\\192.168.9.74\f\ALL REPORT\01. JANUARI 2024\ALL REPORT GABUNGAN",
    # archive_file=r"c:/Users/DELL/Desktop/ReportApp/data/Archive/New AWB Smartfren.csv",
    # output_file=r"c:/Users/DELL/Desktop/ReportApp/data/New AWB Smartfren.csv"
    # )

    print("Menggabungkan semua file CSV Open AWB...")
    merge_csv_files(
        input_folder="c:/Users/DELL/Desktop/ReportApp/data",
        output_folder="c:/Users/DELL/Desktop/ReportApp/data",
        filename_prefix="Update Ryan"
    )

    print("Mengunggah dan mengunduh data dari Apex...")
    tanggal_str = datetime.today().strftime("%d%m%y")
    file_name = f"Update Ryan {tanggal_str}.csv"
    file_dir = r"C:\Users\DELL\Desktop\ReportApp\data"
    download_dir = r"C:\Users\DELL\Desktop\ReportApp\data\apex_downloads"
    file_path = os.path.join(file_dir, file_name)
    base_file_name = os.path.splitext(file_name)[0]

    apex_url = "http://10.18.2.35:7777/apex/f?p=105:45:::NO::P45_PARAM_TYPE:A"
    username = "ccc.support4@jne.co.id"
    password = "123"

    print("Memulai proses upload-download APEX...")
    driver = start_driver(download_dir)

    try:
        # Step 1: Login
        login_to_apex(driver, apex_url, username, password)

        # Step 2: Upload
        upload_file(driver, file_path, file_name)

        # Step 3: Klik Download
        click_download_link(driver, file_name)

        # Step 4: Cek hasil download
        downloaded_file = wait_for_download(download_dir, base_file_name)

        if downloaded_file:
            print("Siap lanjut proses lain (misalnya: request data by account)...")
            # Panggil fungsi lain di sini, masih bisa pakai `driver`
            # contoh:
            # request_by_account(driver, downloaded_file)
        else:
            print("Gagal melanjutkan karena file tidak ditemukan.")

    finally:
        driver.quit()


    print("Semua proses selesai.")

if __name__ == "__main__":
    main()
