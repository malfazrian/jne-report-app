import pandas as pd
import os
import datetime

def extract_new_awb_smartfren(
    base_folder,
    archive_file,
    output_file,
    group_name="DISTRIBUSI SENTRA JAYA"
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
            except:
                use_manual_input = True
        else:
            print("File archive tidak ditemukan.")
            use_manual_input = True

        if use_manual_input:
            start_date = today - datetime.timedelta(days=1)
            end_date = today - datetime.timedelta(days=1)
        else:
            end_date = today - datetime.timedelta(days=1)

        # Cegah start_date lebih baru dari end_date
        if start_date > end_date:
            print(f"Start date ({start_date.date()}) > end date ({end_date.date()}), disesuaikan.")
            start_date = end_date

        # Siapkan path file
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
            print("File tidak ditemukan di kedua path:")
            for path in file_paths:
                print(f"Search {path}")
            return

        # Baca file Excel
        df = pd.read_excel(file_path, sheet_name="All Gab Aduy", dtype=str)
        df.columns = df.columns.str.strip()

        required_columns = {"AWB", "TGL_ENTRY", "GROUP_CUST"}
        if not required_columns.issubset(df.columns):
            print("Kolom yang dibutuhkan tidak lengkap dalam file.")
            return

        df["GROUP_CUST"] = df["GROUP_CUST"].astype(str).str.strip().str.upper()

        try:
            df["TGL_ENTRY"] = pd.to_datetime(df["TGL_ENTRY"], errors="coerce")
        except:
            df["TGL_ENTRY"] = pd.to_numeric(df["TGL_ENTRY"], errors="coerce")
            df["TGL_ENTRY"] = pd.to_datetime(df["TGL_ENTRY"], origin="1899-12-30", unit="d")

        df["TGL_ENTRY"] = df["TGL_ENTRY"].dt.date

        # Filter data berdasarkan tanggal & group
        df_filtered = df[
            (df["TGL_ENTRY"] >= start_date.date()) &
            (df["TGL_ENTRY"] <= end_date.date()) &
            (df["GROUP_CUST"] == group_name)
        ][["AWB"]]

        if df_filtered.empty:
            print("Tidak ada data yang memenuhi kriteria.")
            return

        df_filtered["AWB"] = "'" + df_filtered["AWB"].astype(str)
        df_filtered.to_csv(output_file, index=False, header=False)
        print(f"New AWB Smartfren berhasil disimpan: {output_file}")

    except Exception as e:
        print(f"Terjadi kesalahan: {e}")
