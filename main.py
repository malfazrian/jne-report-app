# main.py
from workflows.ops import get_task_paths, get_file_master_from_open_awb_tasks, preprocess_open_awb, append_new_awb, process_rt_awb
from workflows.file_ops import backup_open_awb_files, clear_folder_of_csvs, clear_folder, merge_csv_files, extract_zip_with_password
from workflows.data_ops import remove_duplicates_by_column, remove_columns_from_file
from workflows.extractors.open_awb_extractor import run_open_awb_extraction
from workflows.extractors.new_awb_extractor import run_new_awb_extraction
from workflows.extractors.awb_rt_extractor import run_rt_awb_extraction
from workflows.extractors.apex_data_extractor import process_apex_upload_and_request
from workflows.email_ops import buka_thunderbird, refresh_inbox, find_latest_matching_email, save_attachments_danamon, extract_date_from_subject
from tasks.ryan_tasks import apex_config, open_awb_tasks, new_awb_tasks, rt_awb_tasks, list_customer_ids
from workflows.tracker import TaskTableTracker
from workflows.danamon_ops import process_pickup_data

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
    account_ids = ["'81635700", "'81635701"]

    task_paths = get_task_paths("c:/Users/DELL/Desktop/ReportApp/data/tracker/tracker_summary.csv")
    updated_open_awb_danamon_path = task_paths.get("Open Danamon")
    updated_rt_danamon_path = task_paths.get("RT Danamon")
    new_awb_danamon_00_path = task_paths.get("New Danamon 00")
    new_awb_danamon_01_path = task_paths.get("New Danamon 01")

    file_master = get_file_master_from_open_awb_tasks(open_awb_tasks, "Open Danamon")

    # === Proses Open AWB ===
    if updated_open_awb_danamon_path and file_master:
        preprocess_open_awb(file_master, updated_open_awb_danamon_path, columns_to_remove, account_ids)
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
        process_rt_awb(updated_rt_danamon_path)
        tracker.set_praprocess("RT Danamon", True)
    else:
        print("Path RT Danamon tidak ditemukan di tracker_summary.csv")

    print("Semua proses praprocess Danamon selesai.")

    # === Proses Ambil Data OS dan Pickup dari Email ===
    print("Membuka Thunderbird dan menyegarkan inbox...")
    buka_thunderbird()
    refresh_inbox("D:/email/Email TB/outlook.office365.com/Inbox.sbd/DANAMON", timeout=10)

    print("Mencari email dengan subjek 'OS CARD JNE'...")
    email_os = find_latest_matching_email(
        file_path="D:/email/Email TB/outlook.office365.com/Inbox.sbd/DANAMON",
        subject_prefix="OS CARD JNE",
        max_emails=10)
    if email_os:
        save_attachments_danamon(email_os, save_dir="d:/RYAN/1. References/Danamon")
        print("Email OS ditemukan dan attachment disimpan.")

    print("Mencari email dengan subjek 'DATA PICKUP JNE'...")
    email_pickup = find_latest_matching_email(
        file_path="D:/email/Email TB/outlook.office365.com/Inbox.sbd/DANAMON",
        subject_prefix="DATA PICKUP JNE",
        max_emails=10)
    if email_pickup:
        save_attachments_danamon(email_pickup, save_dir="d:/RYAN/1. References/Danamon/Ref Pickup/Mentah Txt")
        print("Email Pickup ditemukan dan attachment disimpan.")
        pickup_date = extract_date_from_subject(email_pickup.get("Subject", ""))
        mentah_folder = "d:/RYAN/1. References/Danamon/Ref Pickup/Mentah Txt"
        if pickup_date:
            print("Memproses data pickup dari folder 'Mentah Txt'...")
            process_pickup_data(mentah_folder, subject_date=pickup_date)
        else:
            print("Gagal membaca tanggal dari subject.")

    # tracker.summary()
    print("Semua proses selesai.")

if __name__ == "__main__":
    main()