import os
import glob
import pandas as pd

def extract_open_awb(
    input_path,
    output_path,
    file_type="csv",                # "csv" or "excel"
    status_column="STATUS_POD",
    awb_column="AWB",
    exclude_statuses=None,          # e.g. ["SUCCESS", "RETURN SHIPPER", "DELIVERED"]
    header_row=0,
    quote_awb=False,                # untuk kasih tanda kutip di depan AWB
    process_all_sheets=False        # untuk baca semua sheet dalam file Excel
):
    exclude_statuses = exclude_statuses or ["SUCCESS", "RETURN SHIPPER"]

    try:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        awb_list = []

        if file_type == "csv":
            try:
                df = pd.read_csv(input_path, encoding="utf-8", header=header_row)
            except UnicodeDecodeError:
                df = pd.read_csv(input_path, encoding="latin1", header=header_row)

            print(f"✅ CSV dibaca: {os.path.basename(input_path)}")
            awb_list = [df]

        elif file_type == "excel":
            file_list = glob.glob(os.path.join(input_path, "*.xlsx"))
            for file in file_list:
                try:
                    if process_all_sheets:
                        xls = pd.ExcelFile(file)
                        for sheet_name in xls.sheet_names:
                            df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row, dtype=str)
                            print(f"✅ {os.path.basename(file)} - Sheet: {sheet_name} dibaca.")
                            awb_list.append(df)
                    else:
                        df = pd.read_excel(file, header=header_row, dtype=str)
                        print(f"✅ Excel dibaca: {os.path.basename(file)}")
                        awb_list.append(df)
                except Exception as e:
                    print(f"❌ Gagal membaca {file}: {e}")
        else:
            raise ValueError("❌ Jenis file tidak didukung: gunakan 'csv' atau 'excel'.")

        combined = []
        for df in awb_list:
            df.columns = df.columns.str.strip()

            if {awb_column, status_column}.issubset(df.columns):
                df[status_column] = df[status_column].astype(str).str.strip().str.upper()
                df_filtered = df[~df[status_column].isin([s.upper() for s in exclude_statuses])][[awb_column]]
                if not df_filtered.empty:
                    combined.append(df_filtered)
            else:
                print(f"⚠️ Kolom {awb_column} dan {status_column} tidak ditemukan!")

        if combined:
            result_df = pd.concat(combined, ignore_index=True)

            if quote_awb:
                result_df[awb_column] = "'" + result_df[awb_column].astype(str)

            result_df.to_csv(output_path, index=False, header=False, encoding="utf-8-sig")
            print(f"\n✅ Open AWB disimpan di: {output_path}")
        else:
            print("\n⚠️ Tidak ada data AWB yang valid ditemukan.")

    except Exception as e:
        print(f"❌ Gagal mengekstrak open AWB: {e}")
