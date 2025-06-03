# main.py

from workflows.backup_open_awb import backup_open_awb_files
from workflows.extractors.open_awb_generic_extractor import extract_open_awb
from workflows.extractors.awb_rt_extractor import extract_rt_awb
from workflows.extractors.generali_open_extractor import extract_open_awb_generali

def main():
    print("Memulai ReportApp..")
    
    # 1 Backup file open AWB ke Archive
    print("Menjalankan proses backup file open AWB...")
    backup_open_awb_files()
    
    # 2 Ekstrak open AWB Danamon
    extract_open_awb(
    input_path="D:/RYAN/3. Reports/Danamon/Raw Report.csv",
    output_path="c:/Users/DELL/Desktop/ReportApp/data/Open_AWB_Danamon.csv",
    file_type="csv",
    exclude_statuses=["SUCCESS", "RETURN SHIPPER"]
    )

    # 3 Ekstrak RT AWB Danamon
    print("Mengekstrak RT AWB dari Danamon...")
    extract_rt_awb(
        input_path="D:/RYAN/3. Reports/Danamon/Raw Report.csv",
        output_path="c:/Users/DELL/Desktop/ReportApp/data/RT_AWB_Danamon.csv"
    )

    # 4 Ekstrak open AWB Generali
    print(" Mengekstrak Open AWB dari Generali...")
    extract_open_awb_generali(
        input_folder="D:/RYAN/3. Reports/Generali",
        output_path="c:/Users/DELL/Desktop/ReportApp/data/Open_AWB_Generali.csv"
    )

    print("Mengekstrak Open AWB Kalog...")
    extract_open_awb(
        input_path="D:/RYAN/3. Reports/Kalog",
        output_path="c:/Users/DELL/Desktop/ReportApp/data/Open_AWB_Kalog.csv",
        file_type="excel",
        exclude_statuses=["SUCCESS", "RETURN SHIPPER", "DELIVERED"]
    )
    # from workflows.process_x import process_x
    # process_x()

    print("Semua proses selesai.")

if __name__ == "__main__":
    main()
