import pandas as pd
import re
import string
from typing import List, Optional

def safe_read_csv(path, **kwargs):
    try:
        return pd.read_csv(path, encoding="utf-8", **kwargs)
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="latin1", **kwargs)

def clean_name(text, badwords=None, gelar=None):
    if pd.isna(text):
        return ""

    # Lowercase dan hapus tanda baca
    text = str(text).lower()
    text = re.sub(f"[{re.escape(string.punctuation)}]", "", text)

    # Hapus gelar
    if gelar:
        for g in gelar:
            text = re.sub(rf"\b{re.escape(g.lower())}\b", "", text)

    # Hapus badwords
    if badwords:
        for bad in badwords:
            text = re.sub(rf"\b{re.escape(bad.lower())}\b", "", text)

    # Hapus spasi berlebih
    text = re.sub(r"\s+", " ", text).strip()
    return text

def bersihkan_data_excel(filepath, output_filepath):
    # Load Excel
    df = pd.read_excel(filepath)

    badwords = [
        "keluarga", "kluarga", "kotak", "paket", "tetangga", "teman", "security", "monitoring",
        "toko", "scrty", "lobby", "warung", "kantor", "perusahaan", "pabrik", "cabang", "agen",
        "online", "shop", "keranjang", "keuskupan", "gereja", "masjid", "pura", "vihara", "tb",
        "klenteng", "rumah", "teras", "ibadah", "sakit", "rs", "umum", "swasta", "pemerintah",
        "pos", "indonesia", "kjpp", "kpp", "rekan", "tidak", "tidak mau", "tidak mau di", "menolak", "foto", "di foto", "d foto", "tempat", "mailbox", "jaga", "satpam", "barang", "kiriman", "ruang", "qq", "ybs", "yang bersangkutan", "titip", "titip di", "titip ke", "depan", "belakang", "samping", "atas", "bawah", "kiri", "kanan", "tempat", "lokasi", "alamat", "besmen", "basement", "lantai", "ruang", "area", "kamar", "kamar mandi", "wc", "toilet", "dapur", "kitchen", "ruang tamu", "ruang keluarga", "ruang kerja", "ruang meeting", "ruang rapat", "ruang konferensi", "ruang tunggu", "lobby", "garasi", "parkir", "halaman", "taman", "kebun", "teras belakang", "teras depan", "doc"]

    gelar = ["bpk", "ibu", "dr", "h", "hj", "sdr", "sdri", "bapak", "ibuk", "tuan", "nyonya", "nenek", "kakek", "tante",
        "saudara", "saudari", "pak", "cv", "pt", "pegawai", "karyawan", "staff", "staf", "stafff"] 

    # Bersihkan kolom RECEIVED/REASON
    cleaned = []
    for idx, row in df.iterrows():
        received = clean_name(row.get("RECEIVED/REASON", ""), badwords, gelar)

        if not received:
            # Jika kosong, ambil dari CONSIGNEE_NAME
            received = clean_name(row.get("CONSIGNEE_NAME", ""))

        cleaned.append(received)

    df["CLEAN_RECEIVER"] = cleaned

    # Simpan ke file baru
    df.to_excel(output_filepath, index=False)
    print(f"File berhasil disimpan: {output_filepath}")

def remove_rows_by_prefix(
    input_path: str,
    filter_column: str,
    starts_with: str,
    output_path: str,
    output_format: str = "csv"
):
    """
    Menghapus baris yang kolomnya dimulai dengan prefix tertentu, lalu simpan hasilnya.

    Parameters:
        input_path (str): Path ke file CSV input.
        filter_column (str): Nama kolom untuk pengecekan prefix.
        starts_with (str): Nilai awalan string yang ingin dihapus.
        output_path (str): Path file output.
        output_format (str): Format output: 'csv' atau 'excel'.

    Returns:
        None
    """
    try:
        # Baca file CSV
        df = safe_read_csv(input_path)

        # Pastikan kolom yang diminta ada
        if filter_column not in df.columns:
            raise ValueError(f"Kolom '{filter_column}' tidak ditemukan di data.")

        # Buang baris yang kolomnya dimulai dengan string tertentu
        cleaned_df = df[~df[filter_column].astype(str).str.startswith(starts_with)]

        # Simpan ke format yang diminta
        if output_format.lower() == "csv":
            cleaned_df.to_csv(output_path, index=False)
        elif output_format.lower() == "excel":
            cleaned_df.to_excel(output_path, index=False)
        else:
            raise ValueError("Format output harus 'csv' atau 'excel'.")

        print(f"File berhasil disimpan setelah menghapus baris: {output_path}")

    except Exception as e:
        print(f"Terjadi kesalahan: {e}")

def remove_rows_by_value_exclusion(
    input_path: str,
    filter_column: str,
    allowed_values: list,
    output_path: str,
    output_format: str = "csv"
):
    """
    Menghapus baris yang kolom tertentu nilainya TIDAK termasuk allowed_values.

    Parameters:
        input_path (str): Path file input (CSV/Excel).
        filter_column (str): Nama kolom yang ingin difilter.
        allowed_values (list): Nilai-nilai yang boleh dipertahankan.
        output_path (str): Path untuk menyimpan hasil output.
        output_format (str): Format output: 'csv' atau 'excel'.

    Returns:
        None
    """
    try:
        # Deteksi format input berdasarkan ekstensi
        if input_path.lower().endswith(".csv"):
            df = safe_read_csv(input_path)
        elif input_path.lower().endswith((".xls", ".xlsx")):
            df = pd.read_excel(input_path)
        else:
            raise ValueError("Format file input harus .csv, .xls, atau .xlsx")

        # Validasi kolom
        if filter_column not in df.columns:
            raise ValueError(f"Kolom '{filter_column}' tidak ditemukan.")

        # Hapus baris yang tidak termasuk allowed_values
        filtered_df = df[df[filter_column].astype(str).str.strip().isin(allowed_values)]

        # Simpan hasil
        if output_format.lower() == "csv":
            filtered_df.to_csv(output_path, index=False)
        elif output_format.lower() == "excel":
            filtered_df.to_excel(output_path, index=False)
        else:
            raise ValueError("Format output harus 'csv' atau 'excel'.")

        print(f"Hasil berhasil disimpan ke: {output_path}")

    except Exception as e:
        print(f"Terjadi kesalahan: {e}")

def append_selected_columns_to_master(
    apex_input_path: str,
    master_file_path: str,
    columns_to_exclude: Optional[List[str]] = None,
    allowed_account_ids: Optional[List[str]] = None,
    output_format: str = "csv"
):
    """
    Menambahkan data dari file Apex ke file master, dengan filter ID_ACCOUNT dan
    pengecualian kolom tertentu, tanpa mengubah file Apex aslinya.

    Parameters:
        apex_input_path (str): Path file hasil Apex.
        master_file_path (str): Path file master yang akan diperbarui.
        columns_to_exclude (list): List nama kolom yang ingin dihapus sebelum digabung ke master.
        allowed_account_ids (list): Daftar ID_ACCOUNT yang boleh digabungkan.
        output_format (str): Format file master: 'csv' atau 'excel'.

    Returns:
        None
    """
    try:
        # Baca file Apex
        apex_df = (
            safe_read_csv(apex_input_path)
            if apex_input_path.endswith(".csv")
            else pd.read_excel(apex_input_path)
        )

        # Filter berdasarkan allowed_account_ids jika disediakan
        if allowed_account_ids is not None:
            if "ID_ACCOUNT" not in apex_df.columns:
                raise KeyError("Kolom 'ID_ACCOUNT' tidak ditemukan dalam file Apex.")
            apex_df = apex_df[apex_df["ID_ACCOUNT"].astype(str).isin([str(i) for i in allowed_account_ids])]

        # Baca file master (jika ada), atau buat df kosong
        try:
            master_df = (
                safe_read_csv(master_file_path)
                if master_file_path.endswith(".csv")
                else pd.read_excel(master_file_path)
            )
        except FileNotFoundError:
            master_df = pd.DataFrame()

        # Drop kolom yang tidak diperlukan
        columns_to_exclude = columns_to_exclude or []
        apex_filtered = apex_df.drop(columns=[col for col in columns_to_exclude if col in apex_df.columns])

        # Gabungkan
        combined_df = pd.concat([master_df, apex_filtered], ignore_index=True)

        # Simpan hasil
        if output_format.lower() == "csv":
            combined_df.to_csv(master_file_path, index=False)
        elif output_format.lower() == "excel":
            combined_df.to_excel(master_file_path, index=False)
        else:
            raise ValueError("Format output harus 'csv' atau 'excel'.")

        print(f"Data berhasil ditambahkan ke file master: {master_file_path}")

    except Exception as e:
        print(f"Terjadi kesalahan: {e}")
    
def remove_duplicates_by_column(
    input_path: str,
    filter_column: str,
    output_path: str,
    output_format: str = "csv"
):
    """
    Menghapus baris yang duplikat berdasarkan nilai di satu kolom tertentu.

    Parameters:
        input_path (str): Path ke file input (csv/xlsx).
        filter_column (str): Nama kolom yang jadi dasar pengecekan duplikat.
        output_path (str): Path untuk menyimpan hasil bersih.
        output_format (str): Format file output: 'csv' atau 'excel'.

    Returns:
        None
    """
    try:
        # Baca file input
        if input_path.endswith(".csv"):
            df = safe_read_csv(input_path)
        else:
            df = pd.read_excel(input_path)

        # Validasi kolom
        if filter_column not in df.columns:
            raise ValueError(f"Kolom '{filter_column}' tidak ditemukan.")

        # Hapus duplikat berdasarkan kolom tersebut (keep first)
        df_cleaned = df.drop_duplicates(subset=filter_column, keep="first")

        # Simpan hasil
        if output_format.lower() == "csv":
            df_cleaned.to_csv(output_path, index=False)
        elif output_format.lower() == "excel":
            df_cleaned.to_excel(output_path, index=False)
        else:
            raise ValueError("Format output harus 'csv' atau 'excel'.")

        print(f"Duplikat berdasarkan kolom '{filter_column}' berhasil dihapus. Hasil disimpan di: {output_path}")

    except Exception as e:
        print(f"Terjadi kesalahan: {e}")

def remove_columns_from_file(file_path, columns_to_remove):
    df = safe_read_csv(file_path)
    df = df.drop(columns=[col for col in columns_to_remove if col in df.columns])
    df.to_csv(file_path, index=False)
    print(f"Kolom {columns_to_remove} berhasil dihapus dari {file_path}")