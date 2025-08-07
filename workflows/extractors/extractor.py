from workflows.data_ops import (safe_read_csv)

def filter_start_with_and_save_data(
    input_path: str,
    filter_column: str,
    starts_with: str,
    output_path: str,
    output_format: str = "csv"
):
    """
    Memfilter data dari file CSV berdasarkan kolom dan prefix tertentu, lalu menyimpan hasilnya.
    """
    try:
        # Baca file CSV
        df = safe_read_csv(input_path)

        # Pastikan kolom yang diminta ada
        if filter_column not in df.columns:
            raise ValueError(f"Kolom '{filter_column}' tidak ditemukan di data.")

        # Filter baris berdasarkan prefix
        filtered_df = df[df[filter_column].astype(str).str.startswith(starts_with)]

        if len(filtered_df) == 0:
            print("Tidak ada baris yang cocok, file tidak disimpan.")
            return False

        # Simpan hasil sesuai format yang diminta
        if output_format.lower() == "csv":
            filtered_df.to_csv(output_path, index=False)
        elif output_format.lower() == "excel":
            filtered_df.to_excel(output_path, index=False)
        else:
            raise ValueError("Format output harus 'csv' atau 'excel'.")

        print(f"Hasil disimpan ke: {output_path}")

        return True

    except Exception as e:
        print(f"Terjadi kesalahan: {e}")
        return False