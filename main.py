# main.py

from workflows.backup_open_awb import backup_open_awb_files
from workflows.extractors.open_awb_extractor import extract_open_awb
from workflows.extractors.awb_rt_extractor import extract_rt_awb

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

    print("Semua proses selesai.")

if __name__ == "__main__":
    main()
