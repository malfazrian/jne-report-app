# main.py
from workflows.ops import get_task_paths, get_file_master_from_open_awb_tasks, preprocess_open_awb, append_new_awb, process_rt_awb
from workflows.file_ops import backup_open_awb_files, clear_folder_of_csvs, merge_csv_files
from workflows.data_ops import remove_duplicates_by_column, remove_columns_from_file
from workflows.extractors.open_awb_extractor import run_open_awb_extraction
from workflows.extractors.new_awb_extractor import run_new_awb_extraction
from workflows.extractors.awb_rt_extractor import run_rt_awb_extraction
from workflows.extractors.apex_data_extractor import process_apex_upload_and_request
from tasks.ryan_tasks import open_awb_tasks, new_awb_tasks, rt_awb_tasks, list_customer_ids, apex_config
from workflows.tracker import TaskTableTracker

tracker = TaskTableTracker(open_awb_tasks + new_awb_tasks + rt_awb_tasks + list_customer_ids)

def main():
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
        filename_prefix="Update Ryand"
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
    if success:
        print("Proses upload dan unduh APEX berhasil.")
    else:
        print("Proses upload dan unduh APEX gagal.")

    tracker.summary()

    # === Proses Laporan Danamon ===
    print("=== Mulai Proses Praprocess Laporan Danamon dari hasil APEX ===")

    columns_to_remove = [
        "HO_OFFICE_NO", "HO_OFFICE_DATE", "LATEST_SM_ORIGIN", "LATEST_SM_DEST", "LATEST_SM_NO", "LATEST_SM_DATE",
        "1ST_PREVIOUS_SM_ORIGIN", "1ST_PREVIOUS_SM_DEST", "1ST_PREVIOUS_SM_NO", "1ST_PREVIOUS_SM_DATE", 
        "2ND_PREVIOUS_SM_ORIGIN", "2ND_PREVIOUS_SM_DEST", "2ND_PREVIOUS_SM_NO", "2ND_PREVIOUS_SM_DATE", "STATUS_WEB"
    ]

    task_paths = get_task_paths("c:/Users/DELL/Desktop/ReportApp/data/tracker/tracker_summary.csv")
    updated_open_awb_danamon_path = task_paths.get("Open Danamon")
    updated_rt_danamon_path = task_paths.get("RT Danamon")
    new_awb_danamon_00_path = task_paths.get("New Danamon 00")
    new_awb_danamon_01_path = task_paths.get("New Danamon 01")

    file_master = get_file_master_from_open_awb_tasks(open_awb_tasks, "Open Danamon")

    # === Proses Open AWB ===
    if updated_open_awb_danamon_path and file_master:
        preprocess_open_awb(file_master, updated_open_awb_danamon_path, columns_to_remove)
        tracker.set_praprocess("Open Danamon", True)

    # === Gabungkan New AWB ===
    if new_awb_danamon_00_path:
        append_new_awb(new_awb_danamon_00_path, file_master, "New Danamon 00", columns_to_remove)
        tracker.set_praprocess("New Danamon 00", True)
    else:
        print("Path New Danamon 00 tidak ditemukan di tracker_summary.csv")

    if new_awb_danamon_01_path:
        append_new_awb(new_awb_danamon_01_path, file_master, "New Danamon 01", columns_to_remove)
        tracker.set_praprocess("New Danamon 01", True)
    else:
        print("Path New Danamon 01 tidak ditemukan di tracker_summary.csv")

    remove_columns_from_file(file_master, columns_to_remove)

    # === Hapus Duplikat ===
    print("Hapus duplikat berdasarkan AWB di file master...")
    remove_duplicates_by_column(
        input_path=file_master,
        filter_column="AWB",
        output_path=file_master,
        output_format="csv"
    )

    # === Proses RT AWB ===
    if updated_rt_danamon_path:
        process_rt_awb(updated_rt_danamon_path, tracker)
        tracker.set_praprocess("RT Danamon", True)
    else:
        print("Path RT Danamon tidak ditemukan di tracker_summary.csv")

    print("Semua proses praprocess Danamon selesai.")

    tracker.summary()
    print("Semua proses selesai.")

if __name__ == "__main__":
    main()