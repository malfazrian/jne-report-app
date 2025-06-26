import os
import pandas as pd

def extract_rt_awb(desc, input_path, output_path, tracker=None):
    try:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        try:
            df = pd.read_csv(input_path, encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(input_path, encoding="latin1")

        print(f"File {desc} berhasil dibaca!")

        df.columns = df.columns.str.strip()

        required_columns = {"CONNOTE_RETURN_RT", "STATUS_POD"}
        if not required_columns.issubset(df.columns):
            detail = f"Kolom tidak lengkap. Kolom tersedia: {df.columns.tolist()}"
            print(detail)
            if tracker:
                tracker.set_preprocess(desc, False)
            return

        df["STATUS_POD"] = df["STATUS_POD"].astype(str).str.strip().str.upper()
        df_filtered = df[df["STATUS_POD"].isin(["RU SHIPPER/ORIGIN"])][["CONNOTE_RETURN_RT"]]

        if df_filtered.empty:
            detail = f"Tidak ada data RT AWB dengan status 'RU SHIPPER/ORIGIN' ditemukan."
            print(detail)
            if tracker:
                tracker.set_preprocess(desc, False)
            return

        df_filtered.to_csv(output_path, index=False, header=False, encoding="utf-8-sig")
        print(f"RT AWB {desc} disimpan ke: {output_path}")
        if tracker:
            tracker.set_preprocess(desc, True)

    except Exception as e:
        detail = f"Gagal mengekstrak RT AWB: {e}"
        print(detail)
        if tracker:
            tracker.set_preprocess(desc, False)

def run_rt_awb_extraction(tasks: list, tracker=None):
    for task in tasks:
        print(f"Memproses RT AWB {task['desc']}...")
        extract_rt_awb(
            desc=task["desc"],
            input_path=task["input_path"],
            output_path=task["output_path"],
            tracker=tracker
        )
        print(f"Proses {task['desc']} selesai.\n")
