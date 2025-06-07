# main.py
from workflows.file_ops import backup_open_awb_files, clear_folder_of_csvs, merge_csv_files
from workflows.extractors.open_awb_extractor import run_open_awb_extraction
from workflows.extractors.new_awb_extractor import run_new_awb_extraction
from workflows.extractors.awb_rt_extractor import run_rt_awb_extraction
from workflows.extractors.apex_data_extractor import process_apex_upload_and_request
from tasks.ryan_tasks import open_awb_tasks, new_awb_tasks, rt_awb_tasks, list_customer_ids, apex_config


def main():
    print("Memulai ReportApp..")
    
    # 1 Backup file sebelumnya ke Archive
    print("Menjalankan proses backup file sebelumnya...")
    backup_open_awb_files()
    
    # 2 Ekstrak Open AWB
    print("Mengekstrak Open AWB dari berbagai sumber...")
    run_open_awb_extraction(open_awb_tasks)

    # 3 Ekstrak RT AWB
    print("Mengekstrak RT AWB dari berbagai sumber...")
    run_rt_awb_extraction(rt_awb_tasks)

    # # 4 Ekstrak New AWB
    # print("Mengekstrak New AWB dari berbagai sumber...")
    # run_new_awb_extraction(new_awb_tasks)

    # 5 Merge file hasil ekstraksi AWB
    print("Menggabungkan semua file CSV Open AWB...")
    file = merge_csv_files(
        input_folder="c:/Users/DELL/Desktop/ReportApp/data",
        output_folder="c:/Users/DELL/Desktop/ReportApp/data",
        filename_prefix="Update Ryan"
    )

    # 6 Hapus file CSV lama di folder apex_downloads
    clear_folder_of_csvs("c:/Users/DELL/Desktop/ReportApp/data/apex_downloads")

    # 7 Proses Upload APEX
    print("Mengunggah dan mengunduh data dari Apex...")
    success = process_apex_upload_and_request(
        file_name=file,
        base_file_name="Update Ryan",
        file_path=apex_config["upload_endpoint"],
        download_dir=apex_config["download_endpoint"],
        list_customer_ids=list_customer_ids
    )
    if success:
        print("Proses upload dan unduh APEX berhasil.")
    else:
        print("Proses upload dan unduh APEX gagal.")
    print("Semua proses selesai.")

if __name__ == "__main__":
    main()
