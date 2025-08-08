import os
import sys
from workflows.logger import DualLogger
from workflows.ops import get_task_paths
from workflows.ops import get_file_master_from_open_awb_tasks, preprocess_open_awb
from workflows.extractors.open_awb_extractor import run_open_awb_extraction
from workflows.extractors.apex_data_extractor import process_apex_upload_and_request
from workflows.file_ops import backup_open_awb_files, clear_folder_of_csvs, merge_csv_files, refresh_excel_workbooks
from workflows.data_ops import remove_duplicates_by_column, remove_columns_from_file
from workflows.danamon_ops import convert_to_fixed_width, update_submitted_report
from workflows.whatsapp_ops import send_all_files
from workflows.tracker import TaskTableTracker
from tasks.danamon_tasks import apex_config, open_awb_tasks, new_awb_tasks, rt_awb_tasks, list_customer_ids

def main(tracker, task_paths):
    print("=== Mulai Proses Praprocess Laporan Danamon dari hasil APEX ===")
    sys.stdout = sys.stderr = DualLogger("data/tracker/log.txt")
    print("Memulai ReportApp..")
    
    print("Menjalankan proses backup file sebelumnya...")
    backup_open_awb_files()
    
    print("Mengekstrak Open AWB dari berbagai sumber...")
    run_open_awb_extraction(open_awb_tasks, tracker=tracker)

    print("Menggabungkan semua file CSV Open AWB...")
    file = merge_csv_files(
        input_folder="c:/Users/DELL/Desktop/ReportApp/data",
        output_folder="c:/Users/DELL/Desktop/ReportApp/data",
        filename_prefix="Update Danamon"
    )

    clear_folder_of_csvs("c:/Users/DELL/Desktop/ReportApp/data/apex_downloads")

    print("Mengunggah dan mengunduh data dari Apex...")
    success = process_apex_upload_and_request(
        file_name=file,
        base_file_name="Update Danamon",
        file_path=apex_config["upload_endpoint"],
        download_dir=apex_config["download_endpoint"],
        list_customer_ids=list_customer_ids,
        tracker=tracker,
        open_awb_tasks=open_awb_tasks,
        new_awb_tasks=new_awb_tasks,
        rt_awb_tasks=rt_awb_tasks
    )
    if success:
        print("Proses upload dan unduh APEX berhasil.")
    else:
        print("Proses upload dan unduh APEX gagal.")

    tracker.summary(print_summary=False)
    task_paths = get_task_paths("c:/Users/DELL/Desktop/ReportApp/data/tracker/tracker_summary.csv")

    columns_to_remove = [
        "HO_OFFICE_NO", "HO_OFFICE_DATE", "LATEST_SM_ORIGIN", "LATEST_SM_DEST", "LATEST_SM_NO", "LATEST_SM_DATE",
        "1ST_PREVIOUS_SM_ORIGIN", "1ST_PREVIOUS_SM_DEST", "1ST_PREVIOUS_SM_NO", "1ST_PREVIOUS_SM_DATE", 
        "2ND_PREVIOUS_SM_ORIGIN", "2ND_PREVIOUS_SM_DEST", "2ND_PREVIOUS_SM_NO", "2ND_PREVIOUS_SM_DATE", "STATUS_WEB"
    ]
    account_ids = ["'81635700", "'81635701"]

    updated_open_awb_danamon_path = task_paths.get("Open Danamon")

    file_master = get_file_master_from_open_awb_tasks(open_awb_tasks, "Open Danamon")

    if updated_open_awb_danamon_path and file_master:
        preprocess_open_awb(file_master, updated_open_awb_danamon_path, columns_to_remove, account_ids)
        tracker.set_praprocess("Open Danamon", True)
    else:
        print("Path Open Danamon tidak ditemukan di tracker_summary.csv")

    remove_columns_from_file(file_master, columns_to_remove)

    print("Hapus duplikat berdasarkan AWB di file master...")
    remove_duplicates_by_column(
        input_path=file_master,
        filter_column="AWB",
        output_path=file_master,
        output_format="csv"
    )

    print("Semua proses praprocess Danamon selesai.")

    refresh_excel_workbooks(["d:\\RYAN\\2. Queries\\Query Danamon.xlsx"])

    sukses_path = convert_to_fixed_width(
        input_path="d:\\RYAN\\2. Queries\\Query Danamon.xlsx",
        output_dir="d:\\RYAN\\3. Reports\\Danamon\\Report Sukses",
        sheet_name="REPORT SUKSES KK"
    )

    if sukses_path:
        update_submitted_report()

    report_list = [
        {
            'contact': 'Zaini',
            'files': [f for f in [sukses_path] if f and os.path.exists(f)]
        }
    ]

    send_all_files(report_list)

if __name__ == "__main__":
    tracker = TaskTableTracker(open_awb_tasks)
    task_paths = get_task_paths("c:/Users/DELL/Desktop/ReportApp/data/tracker/tracker_summary.csv")
    main(tracker, task_paths)