import os
import pandas as pd
from datetime import datetime
from typing import Optional

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

        print(f"✅ Data pickup ditambahkan ke: {output_file}")
