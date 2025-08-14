# main.py 
import sys
import traceback
from datetime import datetime
from workflows.logger import DualLogger
from workflows.ops import get_task_paths
from workflows.file_ops import backup_open_awb_files, clear_folder_of_csvs, merge_csv_files, refresh_excel_workbooks
from workflows.data_ops import append_selected_columns_to_master, overwrite_excel_table_with_csv
from workflows.extractors.open_awb_extractor import run_open_awb_extraction
from workflows.extractors.new_awb_extractor import run_new_awb_extraction
from workflows.extractors.awb_rt_extractor import run_rt_awb_extraction
from workflows.extractors.apex_data_extractor import process_apex_upload_and_request
from tasks.ryan_tasks import apex_config, open_awb_tasks, new_awb_tasks, rt_awb_tasks, list_customer_ids
from workflows.tracker import TaskTableTracker
from workflows.danamon_ops import process_danamon_report
from workflows.generali_ops import process_generali_reports
tracker = TaskTableTracker(open_awb_tasks + new_awb_tasks + rt_awb_tasks + list_customer_ids)

def main():
    sys.stdout = sys.stderr = DualLogger("data/tracker/log.txt")
    print("Memulai ReportApp..")
    
    print("Menjalankan proses backup file sebelumnya...")
    backup_open_awb_files()
    
    print("Mengekstrak Open AWB dari berbagai sumber...")
    run_open_awb_extraction(open_awb_tasks, tracker=tracker)

    print("Mengekstrak RT AWB dari berbagai sumber...")
    run_rt_awb_extraction(rt_awb_tasks, tracker=tracker)

    print("Mengekstrak New AWB dari berbagai sumber...")
    run_new_awb_extraction(new_awb_tasks, tracker=tracker)

    print("Menggabungkan semua file CSV Open AWB...")
    file = merge_csv_files(
        input_folder="c:/Users/DELL/Desktop/ReportApp/data",
        output_folder="c:/Users/DELL/Desktop/ReportApp/data",
        filename_prefix="Update Ryan"
    )

    clear_folder_of_csvs("c:/Users/DELL/Desktop/ReportApp/data/apex_downloads")

    print("Mengunggah dan mengunduh data dari Apex...")
    success = process_apex_upload_and_request(
        file_name=file,
        base_file_name="Update Ryan",
        file_path=apex_config["upload_endpoint"],
        download_dir=apex_config["download_endpoint"],
        list_customer_ids=list_customer_ids,
        tracker=tracker,
        open_awb_tasks=open_awb_tasks,
        new_awb_tasks=new_awb_tasks,
        rt_awb_tasks=rt_awb_tasks
    )
    success = True
    if success:
        print("Proses upload dan unduh APEX berhasil.")

        tracker.summary(print_summary=False)
        task_paths = get_task_paths("c:/Users/DELL/Desktop/ReportApp/data/tracker/tracker_summary.csv")

        # === Proses Laporan Danamon ===
        process_danamon_report(tracker, task_paths)

        # === Pra Proses Laporan Generali, Kalog, LSIN, Smartfren, BCA Logistik===
        print("=== Mulai Praprocess Laporan Generali, Kalog, LSIN & Smartfren ===")
        updated_open_awb_path = (
            task_paths.get("Open Generali") or
            task_paths.get("Open Kalog") or
            task_paths.get("Open LSIN") or
            task_paths.get("Open Smartfren") or
            task_paths.get("Open BCA Logistik")
        )

        # List task name yang ingin diproses
        target_tasks = [
            "New Kalog", "New LSIN", "New BCA Logistik",
            "New Smartfren 10528201", "New Smartfren 80029200", "New Smartfren 80044400",
            "New Smartfren 80044800", "New Smartfren 80044801", "New Smartfren 80055900",
            "New Smartfren 80539301", "New Smartfren 80539308"
        ]
        
        if updated_open_awb_path:
            for task_name in target_tasks:
                apex_path = task_paths.get(task_name)
                if apex_path:
                    append_selected_columns_to_master(
                        master_file_path=updated_open_awb_path,
                        apex_input_path=apex_path,
                    )
                    tracker.set_praprocess(task_name, True)
                            
            overwrite_excel_table_with_csv(
            file_path="d:\\RYAN\\2. Queries\\1. Data Apex Mentah.xlsx",
            csv_path=updated_open_awb_path,
            sheet_name="Sheet1",
            )
            refresh_excel_workbooks([
                "d:\\RYAN\\2. Queries\\1. Data Apex Mentah.xlsx",
                "d:\\RYAN\\2. Queries\\2. Query Update Status Generali.xlsx",
                "d:\\RYAN\\2. Queries\\2. Query Update Status.xlsx",
                "d:\\RYAN\\2. Queries\\3. Query Report Generali.xlsx",
                "d:\\RYAN\\2. Queries\\3. Query Report KALOG.xlsx",
                "d:\\RYAN\\2. Queries\\3. Query Report LSIN.xlsx",
                "d:\\RYAN\\2. Queries\\3. Query Report Smartfren.xlsx"
            ])
            tracker.set_praprocess("Open Generali", True)
            tracker.set_praprocess("Open Kalog", True)
            tracker.set_praprocess("Open LSIN", True)
            tracker.set_praprocess("Open Smartfren", True)

        else:
            print("Path Updated Open AWB tidak ditemukan di tracker_summary.csv")
            return
        
        # === Proses Laporan Generali===
        # print("=== Mulai Proses Laporan Generali ===")
        # process_generali_reports()

        tracker.summary()
        print("Semua proses selesai.")
        
    else:
        print("Proses upload dan unduh APEX gagal.")

if __name__ == "__main__":
    try:
        main()
    except Exception:
        log_file = "error_log.txt"
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Terjadi error:\n")
            traceback.print_exc(file=f)