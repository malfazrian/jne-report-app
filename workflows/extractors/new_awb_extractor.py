import pandas as pd
import os
import datetime

def extract_new_awb(
    desc,
    base_folder,
    archive_file,
    output_file,
    group_name="DISTRIBUSI SENTRA JAYA",
    tracker=None
):
    try:
        today = datetime.datetime.today()
        use_manual_input = False

        # Tentukan start_date
        if os.path.exists(archive_file):
            try:
                modified_timestamp = os.path.getmtime(archive_file)
                start_date = datetime.datetime.fromtimestamp(modified_timestamp).replace(hour=0, minute=0)
                print(f"Start date otomatis dari file sebelumnya: {start_date.strftime('%d-%m-%Y')}")
            except Exception as e:
                use_manual_input = True
                print(f"Gagal baca tanggal dari archive: {e}")
        else:
            print("File archive tidak ditemukan.")
            use_manual_input = True

        end_date = today - datetime.timedelta(days=1)
        if use_manual_input:
            start_date = end_date

        if start_date > end_date:
            print(f"Start date ({start_date.date()}) > end date ({end_date.date()}), disesuaikan.")
            start_date = end_date

        tanggal_str = today.strftime("%d %m %y")
        bulan_inggris = end_date.strftime("%m. %b %y").upper()

        bulan_map = {
            "JAN": "JANUARI", "FEB": "FEBRUARI", "MAR": "MARET", "APR": "APRIL",
            "MAY": "MEI", "JUN": "JUNI", "JUL": "JULI", "AUG": "AGUSTUS",
            "SEP": "SEPTEMBER", "OCT": "OKTOBER", "NOV": "NOVEMBER", "DEC": "DESEMBER"
        }

        bulan_singkat = bulan_inggris.split(". ")[1].split(" ")[0]
        bulan_indonesia = bulan_inggris.replace(bulan_singkat, bulan_map[bulan_singkat])

        file_paths = [
            os.path.join(base_folder, bulan_inggris, tanggal_str, "All Gab Aduy.xlsx"),
            os.path.join(base_folder, bulan_indonesia, tanggal_str, "All Gab Aduy.xlsx")
        ]

        file_path = next((p for p in file_paths if os.path.exists(p)), None)
        if not file_path:
            detail = f"File tidak ditemukan:\n- {file_paths[0]}\n- {file_paths[1]}"
            print(detail)
            if tracker:
                tracker.set_preprocess(desc, False)
            return

        df = pd.read_excel(file_path, sheet_name="All Gab Aduy", dtype=str)
        df.columns = df.columns.str.strip()

        required_columns = {"AWB", "TGL_ENTRY", "GROUP_CUST"}
        if not required_columns.issubset(df.columns):
            detail = f"Kolom wajib tidak lengkap: {required_columns}"
            print(detail)
            if tracker:
                tracker.set_preprocess(desc, False)
            return

        df["GROUP_CUST"] = df["GROUP_CUST"].astype(str).str.strip().str.upper()

        try:
            df["TGL_ENTRY"] = pd.to_datetime(df["TGL_ENTRY"], errors="coerce")
        except:
            df["TGL_ENTRY"] = pd.to_numeric(df["TGL_ENTRY"], errors="coerce")
            df["TGL_ENTRY"] = pd.to_datetime(df["TGL_ENTRY"], origin="1899-12-30", unit="d")

        df["TGL_ENTRY"] = df["TGL_ENTRY"].dt.date

        df_filtered = df[
            (df["TGL_ENTRY"] >= start_date.date()) &
            (df["TGL_ENTRY"] <= end_date.date()) &
            (df["GROUP_CUST"] == group_name)
        ][["AWB"]]

        if df_filtered.empty:
            detail = f"Tidak ada data memenuhi kriteria tanggal ({start_date.date()} - {end_date.date()}) dan group '{group_name}'."
            print(detail)
            if tracker:
                tracker.set_preprocess(desc, False)
            return

        df_filtered["AWB"] = "'" + df_filtered["AWB"].astype(str)
        df_filtered.to_csv(output_file, index=False, header=False)
        print(f"New AWB {desc} berhasil disimpan: {output_file}")
        if tracker:
            tracker.set_preprocess(desc, True)

    except Exception as e:
        detail = f"Terjadi kesalahan: {e}"
        print(detail)
        if tracker:
            tracker.set_preprocess(desc, False)

def run_new_awb_extraction(new_awb_tasks: list, tracker=None):
    for task in new_awb_tasks:
        print(f"Mengekstrak New AWB untuk {task['desc']}...")
        extract_new_awb(
            desc=task["desc"],
            base_folder=task["base_folder"],
            archive_file=task["archive_file"],
            output_file=task["output_file"],
            tracker=tracker
        )
