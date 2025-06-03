# main.py

from workflows.backup_open_awb import backup_open_awb_files
from workflows.extractors.open_awb_extractor import extract_open_awb
from workflows.extractors.new_awb_extractor import extract_new_awb_smartfren
from workflows.extractors.awb_rt_extractor import extract_rt_awb
from core.file_ops import merge_csv_files

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

    print("Mengekstrak New AWB Smartfren...")
    extract_new_awb_smartfren(
    base_folder=r"\\192.168.9.74\f\ALL REPORT\01. JANUARI 2024\ALL REPORT GABUNGAN",
    archive_file=r"c:/Users/DELL/Desktop/ReportApp/data/Archive/New AWB Smartfren.csv",
    output_file=r"c:/Users/DELL/Desktop/ReportApp/data/New AWB Smartfren.csv"
    )

    print("Menggabungkan semua file CSV Open AWB...")
    merge_csv_files(
        input_folder="c:/Users/DELL/Desktop/ReportApp/data",
        output_folder="c:/Users/DELL/Desktop/ReportApp/data",
        filename_prefix="Update Ryan"
    )

    print("Semua proses selesai.")

if __name__ == "__main__":
    main()
