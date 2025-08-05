import os
import glob
import shutil
import pandas as pd
import pyzipper
import py7zr
import win32com.client
import time
from datetime import datetime
from pathlib import Path
from pywintypes import com_error
from typing import Union, List

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

def extract_archive_with_password(file_path, extract_to, password=None, delete_after=True):
    """
    Mengekstrak file arsip ZIP atau 7Z, dengan atau tanpa password.

    Parameters:
        file_path (str): Path ke file arsip (ZIP atau 7Z).
        extract_to (str): Folder tujuan ekstrak.
        password (str or None): Password jika diperlukan.
        delete_after (bool): Hapus file arsip setelah ekstraksi.

    Returns:
        list: Daftar file yang berhasil diekstrak.
    """
    extracted_files = []

    if file_path.lower().endswith(".zip"):
        try:
            with pyzipper.AESZipFile(file_path) as zf:
                if password:
                    zf.pwd = password.encode("utf-8")

                for file in zf.namelist():
                    try:
                        zf.extract(member=file, path=extract_to)
                        extracted_files.append(file)
                        print(f"Diekstrak (ZIP): {file}")
                    except RuntimeError as e:
                        print(f"Gagal ekstrak (ZIP) '{file}': {e}")

        except Exception as e:
            print(f"Gagal membuka ZIP: {e}")

    elif file_path.lower().endswith(".7z"):
        try:
            with py7zr.SevenZipFile(file_path, mode='r', password=password) as archive:
                archive.extractall(path=extract_to)
                extracted_files = archive.getnames()
                print(f"Diekstrak (7Z): {extracted_files}")
        except Exception as e:
            print(f"Gagal membuka 7Z: {e}")
    else:
        print("Format arsip tidak dikenali. Hanya .zip dan .7z yang didukung.")

    if delete_after and os.path.exists(file_path):
        try:
            os.remove(file_path)
            print(f"File arsip dihapus: {file_path}")
        except Exception as e:
            print(f"Gagal menghapus file: {e}")

    return extracted_files

def wait_for_query_refresh(wb):
    """Tunggu semua koneksi query hingga selesai refresh (tanpa batas waktu)"""
    print("Menunggu query selesai...")
    while True:
        refreshing = False
        try:
            for connection in wb.Connections:
                try:
                    if connection.OLEDBConnection.Refreshing:
                        refreshing = True
                        break
                except AttributeError:
                    continue
        except com_error as e:
            if e.hresult == -2147418111:  # Call was rejected by callee
                print("Excel sedang sibuk... tunggu sebentar.")
                time.sleep(2)
                continue
            else:
                raise

        if not refreshing:
            break

        time.sleep(2)

def close_all_excel_instances():
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        # Loop semua workbook yang terbuka
        while excel.Workbooks.Count > 0:
            wb = excel.Workbooks(1)
            wb.Close(SaveChanges=False)
        # Tutup Excel kalau tidak ada workbook lagi
        excel.Quit()
    except Exception as e:
        print("Terjadi error:", e)

def refresh_excel_workbooks(target: Union[str, Path, List[Union[str, Path]]]):

    """
    Refresh semua koneksi dalam satu atau banyak file Excel (.xlsx).

    Args:
        target (str | Path | List[str | Path]):
            - Path ke folder (semua .xlsx di-refresh)
            - Atau list path file Excel

    Returns:
        None
    """
    if isinstance(target, (str, Path)):
        target = Path(target)
        if target.is_dir():
            file_list = sorted(target.glob("*.xlsx"))
        elif target.is_file():
            file_list = [target]
        else:
            raise FileNotFoundError(f"Path tidak ditemukan: {target}")
    elif isinstance(target, list):
        file_list = [Path(f) for f in target if Path(f).suffix == ".xlsx"]
    else:
        raise ValueError("Parameter 'target' harus folder, file, atau list file.")

    if not file_list:
        print("Tidak ada file Excel ditemukan untuk diproses.")
        return
    
    close_all_excel_instances()

    print("Membuka Excel...")
    excel = win32com.client.Dispatch("Excel.Application")

    try:
        for excel_file in file_list:
            print(f"\nMembuka file: {excel_file.name}")
            wb = excel.Workbooks.Open(str(excel_file))

            print("Merefresh semua koneksi data...")
            wb.RefreshAll()

            wait_for_query_refresh(wb)

            wb.Save()
            wb.Close(SaveChanges=True)
            print("Refresh selesai dan file disimpan.")

    except Exception as e:
        print(f"Terjadi kesalahan: {e}")

    finally:
        excel.Quit()
        print("\nSemua file telah diproses.")