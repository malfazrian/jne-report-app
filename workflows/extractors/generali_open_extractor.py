import os
import glob
import pandas as pd

def extract_open_awb_generali(input_folder, output_path):
    try:
        file_list = glob.glob(os.path.join(input_folder, "*.xlsx"))
        awb_list = []

        for file in file_list:
            try:
                df = pd.read_excel(file, header=2, dtype=str)
                df.columns = df.columns.str.strip()

                required_columns = {"AWB", "DELIVERY STATUS"}
                if required_columns.issubset(df.columns):
                    df["DELIVERY STATUS"] = df["DELIVERY STATUS"].astype(str).str.strip().str.upper()

                    df_filtered = df[~df["DELIVERY STATUS"].isin(["SUCCESS", "RETURN SHIPPER"])][["AWB"]]
                    if not df_filtered.empty:
                        awb_list.append(df_filtered)

                    print(f"{os.path.basename(file)} diproses.")
                else:
                    print(f"Kolom yang dibutuhkan tidak ditemukan di {os.path.basename(file)}!")

            except Exception as file_error:
                print(f"Error saat memproses {os.path.basename(file)}: {file_error}")

        if awb_list:
            result_df = pd.concat(awb_list, ignore_index=True)
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            result_df.to_csv(output_path, index=False, header=False, encoding="utf-8-sig")
            print(f"Open AWB Generali disimpan di: {output_path}")
        else:
            print("Tidak ada data yang memenuhi kriteria.")

    except Exception as e:
        print(f"Gagal mengekstrak Open AWB Generali: {e}")
