import csv
from workflows.data_ops import remove_rows_by_prefix, remove_rows_by_value_exclusion, append_selected_columns_to_master, remove_duplicates_by_column
from workflows.extractors.extractor import filter_start_with_and_save_data

def get_task_paths(tracker_csv_path):
    task_paths = {}
    with open(tracker_csv_path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            task_paths[row["Task"]] = row["Result Path"]
    return task_paths

def get_file_master_from_open_awb_tasks(open_awb_tasks, task_name):
    for task in open_awb_tasks:
        if task["desc"] == task_name:
            return task["input_path"]
    return None

def preprocess_open_awb(file_master, updated_open_awb_path, columns_to_remove=None):
    print("Proses hapus data Open AWB di file master...")
    remove_rows_by_value_exclusion(
        input_path=file_master,
        filter_column="STATUS_POD",
        allowed_values=["Return Shipper", "Success"],
        output_path=file_master,
        output_format="csv"
    )

    print("Gabungkan Updated Open AWB ke file master...")
    append_selected_columns_to_master(
        apex_input_path=updated_open_awb_path,
        master_file_path=file_master,
        columns_to_exclude=columns_to_remove,
        output_format="csv"
    )

def append_new_awb(new_awb_path, file_master, label, columns_to_remove=None):
    if new_awb_path:
        print(f"Gabungkan {label} ke file master...")
        append_selected_columns_to_master(
            apex_input_path=new_awb_path,
            master_file_path=file_master,
            columns_to_exclude=columns_to_remove,
            output_format="csv"
        )

def process_rt_awb(updated_rt_awb_path, tracker):
    print("Mulai proses ekstraksi data RT AWB Danamon dari hasil APEX...")
    result = filter_start_with_and_save_data(
        input_path=updated_rt_awb_path,
        filter_column="AWB",
        starts_with="'RT",
        output_path="D:/RYAN/1. References/Danamon/RT REPORT.csv",
        output_format="csv"
    )

    if result:
        print("Proses hapus data RT AWB dari hasil APEX...")
        remove_rows_by_prefix(
            input_path=updated_rt_awb_path,
            filter_column="AWB",
            starts_with="'RT",
            output_path=updated_rt_awb_path,
            output_format="csv"
        )
        print("Proses RT AWB Danamon selesai.")