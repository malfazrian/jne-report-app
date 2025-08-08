import os
import pandas as pd
import re
from datetime import datetime
from typing import Optional
from workflows.ops import get_file_master_from_open_awb_tasks, preprocess_open_awb, append_new_awb, process_rt_awb
from workflows.data_ops import remove_duplicates_by_column, remove_columns_from_file
from workflows.file_ops import clear_folder, refresh_excel_workbooks
from workflows.email_ops import buka_thunderbird, refresh_inbox, find_latest_matching_email, save_attachments_danamon, extract_date_from_subject
from workflows.whatsapp_ops import send_all_files
from tasks.ryan_tasks import open_awb_tasks

def process_pickup_data(mentah_folder: str, subject_date: Optional[str] = None):
    if not os.path.exists(mentah_folder):
        print(f"Folder tidak ditemukan: {mentah_folder}")
        return

    all_files = [f for f in os.listdir(mentah_folder) if os.path.isfile(os.path.join(mentah_folder, f))]
    if not all_files:
        print("Tidak ada file yang ditemukan di folder Mentah TXT.")
        return

    # Ambil tanggal pickup dari subject email (format 30/06/25)
    pickup_date = "Unknown"
    if subject_date:
        try:
            dt = datetime.strptime(subject_date, "%d/%m/%y")
            pickup_date = dt.strftime("%Y-%m-%d")
        except Exception as e:
            print(f"Format tanggal subject tidak valid: {subject_date} ({e})")

    col_widths = [15, 30, 40, 40, 40, 40, 21, 9, 11]
    column_names = ["REF", "NAMA", "ADD_1", "ADD_2", "ADD_3", "ADD_4", "KOTA", "KODEPOS", "KET"]
    all_data = []

    for file in all_files:
        input_path = os.path.join(mentah_folder, file)
        try:
            df = pd.read_fwf(input_path, widths=col_widths, names=column_names, dtype=str)
            if df.iloc[0].str.contains("REF", na=False).any():
                df = df.iloc[1:].reset_index(drop=True)

            df_filtered = df[["REF"]].copy()
            df_filtered["REF"] = df_filtered["REF"].str.upper()
            df_filtered["DATE PICKUP"] = pickup_date
            df_filtered["FILENAME"] = file
            df_filtered = df_filtered[["DATE PICKUP", "REF", "FILENAME"]]

            all_data.append(df_filtered)

        except Exception as e:
            print(f"Terjadi kesalahan saat memproses {file}: {e}")

    if not all_data:
        print("Tidak ada data yang bisa diproses.")
        return

    final_df = pd.concat(all_data, ignore_index=True)
    final_df["DATE PICKUP"] = pd.to_datetime(final_df["DATE PICKUP"], errors="coerce")
    final_df = final_df.dropna(subset=["DATE PICKUP"])

    # Simpan ke folder berdasarkan bulan
    output_folder = os.path.abspath(os.path.join(mentah_folder, "..", "Data Pickup"))
    os.makedirs(output_folder, exist_ok=True)

    for month_year, group in final_df.groupby(final_df["DATE PICKUP"].dt.strftime("%B %Y")):
        output_file = os.path.join(output_folder, f"Pickup Danamon {month_year}.csv")
        group = group[["DATE PICKUP", "REF", "FILENAME"]]

        if os.path.exists(output_file):
            group.to_csv(output_file, mode="a", header=False, index=False, encoding="utf-8")
        else:
            group.to_csv(output_file, mode="w", header=True, index=False, encoding="utf-8")

        print(f"Data pickup ditambahkan ke: {output_file}")

def left_align(length, value):
    return f"{str(value)[:length]:<{length}}"

def convert_to_fixed_width(input_path, output_dir, sheet_name=None, output_prefix="REPORT SUKSES KK", skip_header_row=True):
    """
    Mengubah file Excel atau CSV menjadi file fixed-width TXT tanpa header.

    Parameters:
        input_path (str): Path ke file CSV atau Excel.
        output_dir (str): Folder tujuan output TXT.
        sheet_name (str): Nama sheet jika file Excel.
        output_prefix (str): Awalan nama file output.
        skip_header_row (bool): Jika True, lewati baris pertama (biasanya header manual).
    """
    if not os.path.exists(input_path):
        print(f"File tidak ditemukan: {input_path}")
        return None

    try:
        if input_path.endswith((".xls", ".xlsx")):
            df = pd.read_excel(input_path, sheet_name=sheet_name, header=None).fillna('')
        else:
            df = pd.read_csv(input_path, header=None).fillna('')
    except Exception as e:
        print(f"Gagal membaca file: {e}")
        return None

    # Panjang kolom fixed-width (sesuaikan dengan kebutuhanmu)
    widths = (17, 21, 2, 2, 2, 9, 9, 40, 7)

    def left_align(length, value):
        return f"{str(value)[:length]:<{length}}"

    record_line_list = []
    data_rows = df.iloc[1:] if skip_header_row else df  # Skip baris pertama jika diaktifkan

    if data_rows.empty:
        print("Tidak ada data yang bisa dikonversi. File tidak dibuat.")
        return None

    for row in data_rows.itertuples(index=False, name=None):
        baris = "".join(left_align(length, value) for length, value in zip(widths, row))
        record_line_list.append(baris)

    current_date = datetime.today().strftime("%d%m%Y")
    filename = f"{output_prefix} {current_date}.txt"
    output_path = os.path.join(output_dir, filename)

    os.makedirs(output_dir, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(record_line_list))

    print(f"Fixed-width file disimpan di: {output_path}")
    return output_path

def rename_file_date(folder_path: str, subject_date: str):
    """
    Rename file dalam folder berdasarkan subject_date.

    Contoh: "OS CC JNE.xlsx" → "OS CC JNE 020725.xlsx" jika subject_date = "02/07/25"
    """

    if not os.path.isdir(folder_path):
        print(f"Folder OS File tidak ditemukan: {folder_path}")
        return

    # Format tanggal untuk nama file: DDMMYY dari subject_date
    try:
        dt = datetime.strptime(subject_date, "%d/%m/%y")
        suffix = dt.strftime("%d%m%y")  # e.g. "020725"
    except ValueError:
        print(f"Format tanggal tidak valid: {subject_date}")
        return

    for fname in os.listdir(folder_path):
        if fname.startswith("OS CC JNE") and fname.lower().endswith(".xlsx"):
            old_path = os.path.join(folder_path, fname)
            name_part, ext = os.path.splitext(fname)

            # Hapus semua tanggal DDMMYY di akhir nama (hindari duplikat)
            name_part = re.sub(r'(\s\d{6})+$', '', name_part)

            new_name = f"{name_part} {suffix}{ext}"
            new_path = os.path.join(folder_path, new_name)

            os.rename(old_path, new_path)
            return new_path
        
def get_report_pickup_filename_from_pickup_date(pickup_date: str) -> str:
    bulan_id_map = {
        "January": "Januari",
        "February": "Februari",
        "March": "Maret",
        "April": "April",
        "May": "Mei",
        "June": "Juni",
        "July": "Juli",
        "August": "Agustus",
        "September": "September",
        "October": "Oktober",
        "November": "November",
        "December": "Desember"
    }

    try:
        dt = datetime.strptime(pickup_date, "%d/%m/%y")
        nama_bulan_inggris = dt.strftime("%B")  # "July"
        nama_bulan_id = bulan_id_map.get(nama_bulan_inggris, nama_bulan_inggris)
        tahun = dt.strftime("%Y")  # "2025"
        return f"Report Pickup Danamon {nama_bulan_id} {tahun}.xlsx"
    except ValueError:
        print("Gagal parsing tanggal pickup.")
        return None
    
def update_submitted_report():
    """
    Update file CSV dengan data dari Excel.
    Mengambil data dari kolom pertama tanpa header, lalu menambahkan tanggal sekarang.
    """
    # File dan sheet sumber
    excel_file = r"D:\RYAN\2. Queries\Query Danamon.xlsx"
    sheet_name = "REPORT SUKSES KK"

    # File tujuan CSV
    csv_file = r"D:\RYAN\1. References\Danamon\Submitted Report Sukses.csv"

    # Ambil data dari kolom pertama tanpa header (skiprows=1 untuk melewati header)
    data = pd.read_excel(excel_file, sheet_name=sheet_name, usecols=[0], skiprows=1, header=None)

    # Ambil hanya baris dengan data (hapus NaN)
    data = data.dropna()

    # Tambahkan kolom tanggal sekarang dalam format 8/6/2025 (tanpa nol di awal)
    today_str = datetime.now().strftime('%#d/%#m/%Y')
    data['Tanggal'] = today_str

    # Simpan / append ke file CSV
    data.to_csv(csv_file, mode='a', index=False, header=False)

    print("Data report sukses berhasil ditambahkan ke file Update Submitted.")
    
def process_danamon_report(tracker, task_paths):
    print("=== Mulai Proses Praprocess Laporan Danamon dari hasil APEX ===")

    columns_to_remove = [
        "HO_OFFICE_NO", "HO_OFFICE_DATE", "LATEST_SM_ORIGIN", "LATEST_SM_DEST", "LATEST_SM_NO", "LATEST_SM_DATE",
        "1ST_PREVIOUS_SM_ORIGIN", "1ST_PREVIOUS_SM_DEST", "1ST_PREVIOUS_SM_NO", "1ST_PREVIOUS_SM_DATE", 
        "2ND_PREVIOUS_SM_ORIGIN", "2ND_PREVIOUS_SM_DEST", "2ND_PREVIOUS_SM_NO", "2ND_PREVIOUS_SM_DATE", "STATUS_WEB"
    ]
    account_ids = ["'81635700", "'81635701"]

    updated_open_awb_danamon_path = task_paths.get("Open Danamon")
    updated_rt_danamon_path = task_paths.get("RT Danamon")
    new_awb_danamon_00_path = task_paths.get("New Danamon 81635700")
    new_awb_danamon_01_path = task_paths.get("New Danamon 81635701")

    file_master = get_file_master_from_open_awb_tasks(open_awb_tasks, "Open Danamon")

    if updated_open_awb_danamon_path and file_master:
        preprocess_open_awb(file_master, updated_open_awb_danamon_path, columns_to_remove, account_ids)
        tracker.set_praprocess("Open Danamon", True)

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

    print("Hapus duplikat berdasarkan AWB di file master...")
    remove_duplicates_by_column(
        input_path=file_master,
        filter_column="AWB",
        output_path=file_master,
        output_format="csv"
    )

    if updated_rt_danamon_path:
        process_rt_awb(updated_rt_danamon_path)
        tracker.set_praprocess("RT Danamon", True)
    else:
        print("Path RT Danamon tidak ditemukan di tracker_summary.csv")

    print("Semua proses praprocess Danamon selesai.")

    # === Proses Email ===
    email_os = None
    email_pickup = None
    pickup_date = None
    pickup_path = None
    os_path = None
    pu_filename = None
    sukses_path = None

    print("Membuka Thunderbird dan menyegarkan inbox...")
    buka_thunderbird()
    refresh_inbox("D:/email/Email TB/outlook.office365.com/Inbox.sbd/DANAMON", timeout=10*60)

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
        mentah_folder = "d:/RYAN/1. References/Danamon/Ref Pickup/Mentah Txt"
        clear_folder(mentah_folder)
        save_attachments_danamon(email_pickup, save_dir=mentah_folder)
        print("Email Pickup ditemukan dan attachment disimpan.")
        pickup_date = extract_date_from_subject(email_pickup.get("Subject", ""))
        if pickup_date:
            print("Memproses data pickup dari folder 'Mentah Txt'...")
            process_pickup_data(mentah_folder, subject_date=pickup_date)
            os_path = rename_file_date(
                folder_path="d:\\RYAN\\3. Reports\\Danamon",
                subject_date=pickup_date,
            )
            pu_filename = get_report_pickup_filename_from_pickup_date(pickup_date)
        else:
            print("Gagal membaca tanggal dari subject.")

    if pu_filename:
        pickup_path = os.path.join("d:\\RYAN\\3. Reports\\Danamon", pu_filename)
        refresh_excel_workbooks(["d:\\RYAN\\2. Queries\\Query Danamon.xlsx", os_path, pickup_path])
    else:
        print("Tidak bisa membangun nama file Excel dari tanggal pickup.")

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
            'files': [f for f in [r'D:\RYAN\3. Reports\Danamon\Raw Report.csv', pickup_path, os_path, sukses_path] if f and os.path.exists(f)]
        }
    ]

    send_all_files(report_list)