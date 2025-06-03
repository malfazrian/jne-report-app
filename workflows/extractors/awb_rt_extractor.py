import os
import pandas as pd

def extract_rt_awb(input_path, output_path):
    try:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        try:
            df = pd.read_csv(input_path, encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(input_path, encoding="latin1")

        print("File Danamon berhasil dibaca!")

        df.columns = df.columns.str.strip()

        required_columns = {"CONNOTE_RETURN_RT", "STATUS_POD"}
        if not required_columns.issubset(df.columns):
            raise ValueError(f"Kolom tidak lengkap. Kolom tersedia: {df.columns.tolist()}")

        df["STATUS_POD"] = df["STATUS_POD"].astype(str).str.strip().str.upper()

        df_filtered = df[df["STATUS_POD"].isin(["RU SHIPPER/ORIGIN"])][["CONNOTE_RETURN_RT"]]

        df_filtered.to_csv(output_path, index=False, header=False, encoding="utf-8-sig")

        print(f"RT AWB Danamon disimpan ke: {output_path}")

    except Exception as e:
        print(f"Gagal mengekstrak RT AWB Danamon: {e}")
