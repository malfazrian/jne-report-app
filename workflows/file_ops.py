import os
import glob
import shutil
import pandas as pd
import pyzipper
from datetime import datetime

def ensure_folder(path):
    if not os.path.exists(path):
        os.makedirs(path)

def clear_folder_of_csvs(folder):
    for file in os.listdir(folder):
        if file.endswith(".csv"):
            os.remove(os.path.join(folder, file))

def clear_folder(folder_path: str):
    """
    Menghapus semua file di dalam folder tertentu (tidak termasuk subfolder).

    Parameters:
        folder_path (str): Path folder yang ingin dibersihkan.

    Returns:
        None
    """
    try:
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
        print(f"Semua file di folder '{folder_path}' berhasil dihapus.")
    except Exception as e:
        print(f"Gagal menghapus isi folder '{folder_path}': {e}")

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
            return os.path.basename(output_file)
        else:
            print("Tidak ada file CSV untuk digabungkan.")
            return "Tidak ada file untuk digabungkan."

def backup_open_awb_files():
    """Backup file open AWB ke folder Archive dan bersihkan arsip sebelumnya."""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.abspath(os.path.join(current_dir, "..", "data"))
    archive_dir = os.path.join(data_dir, "archive")

    ensure_folder(archive_dir)
    clear_folder_of_csvs(archive_dir)
    move_csvs_to_folder(data_dir, archive_dir)

    print("Backup selesai.")

def extract_zip_with_password(zip_path, extract_to, password, delete_after=True):
    """
    Mengekstrak ZIP yang dikunci password menggunakan pyzipper.

    Parameters:
        zip_path (str): Path file zip.
        extract_to (str): Folder tujuan ekstrak.
        password (str): Password zip (tanpa encode manual).
        delete_after (bool): Jika True, hapus file zip setelah ekstraksi.

    Returns:
        list: Daftar file yang berhasil diekstrak.
    """
    extracted_files = []
    try:
        with pyzipper.AESZipFile(zip_path) as zf:
            zf.pwd = password.encode("utf-8")
            for file in zf.namelist():
                try:
                    zf.extract(member=file, path=extract_to)
                    extracted_files.append(file)
                    print(f"Diekstrak: {file}")
                except RuntimeError as e:
                    print(f"Gagal ekstrak '{file}': {e}")

        if delete_after:
            os.remove(zip_path)
            print(f"File ZIP dihapus: {zip_path}")

        print(f"Selesai ekstrak ke: {extract_to}")
        return extracted_files

    except Exception as e:
        print(f"Gagal ekstrak ZIP: {e}")
        return []