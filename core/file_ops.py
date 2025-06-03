import os
import glob
import shutil
import pandas as pd
from datetime import datetime

def ensure_folder(path):
    if not os.path.exists(path):
        os.makedirs(path)

def clear_folder_of_csvs(folder):
    for file in os.listdir(folder):
        if file.endswith(".csv"):
            os.remove(os.path.join(folder, file))

def move_csvs_to_folder(source_dir, destination_dir):
    print(f"Mencari file CSV di: {source_dir}")
    for file in os.listdir(source_dir):
        if file.endswith(".csv"):
            src = os.path.join(source_dir, file)
            dst = os.path.join(destination_dir, file)
            print(f"Memindahkan {file} ke {destination_dir}")
            shutil.move(src, dst)

def merge_csv_files(
    input_folder: str,
    output_folder: str,
    filename_prefix: str = "Update Ryan"
    ) -> str:
        """
        Menggabungkan semua file CSV dalam folder input dan simpan hasilnya ke output_folder.
        
        Args:
            input_folder (str): Path ke folder sumber CSV.
            output_folder (str): Path ke folder tujuan hasil gabungan.
            filename_prefix (str): Nama depan file output.
        
        Returns:
            str: Path lengkap file output, atau pesan jika tidak ada file.
        """
        today_str = datetime.today().strftime("%d%m%y")
        output_file = os.path.join(output_folder, f"{filename_prefix} {today_str}.csv")

        # Ambil semua file CSV
        csv_files = glob.glob(os.path.join(input_folder, "*.csv"))
        merged_data = []

        print(f"Menggabungkan {len(csv_files)} file CSV dari: {input_folder}")

        for file in csv_files:
            try:
                df = pd.read_csv(file, header=None, usecols=[0], dtype=str)
                merged_data.append(df)
                print(f"{os.path.basename(file)} berhasil dimuat.")
            except Exception as e:
                print(f"Gagal membaca {file}: {e}")

        if merged_data:
            result_df = pd.concat(merged_data, ignore_index=True)
            result_df.drop_duplicates(inplace=True)
            
            os.makedirs(output_folder, exist_ok=True)
            result_df.to_csv(output_file, index=False, header=False)

            print(f"File hasil merge disimpan di:\n{output_file}")
            return output_file
        else:
            print("Tidak ada file CSV untuk digabungkan.")
            return ""
