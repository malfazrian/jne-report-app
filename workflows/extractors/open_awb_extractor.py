import os
import glob
import pandas as pd
import warnings

warnings.simplefilter(action='ignore', category=pd.errors.DtypeWarning)

def extract_open_awb(
    input_path,
    output_path,
    file_type="csv",
    status_column="STATUS_POD",
    awb_column="AWB",
    exclude_statuses=None,
    header_row=0,
    quote_awb=False,
    process_all_sheets=False,
    task_name=None,
    tracker=None
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

            print(f"CSV dibaca: {os.path.basename(input_path)}")
            awb_list = [df]

        elif file_type == "excel":
            file_list = glob.glob(os.path.join(input_path, "*.xlsx"))

            if not file_list:
                print(f"Tidak ada file Excel ditemukan di folder: {input_path}")
                return False

            for file in file_list:
                try:
                    if process_all_sheets:
                        xls = pd.ExcelFile(file)
                        for sheet_name in xls.sheet_names:
                            df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row, dtype=str)
                            print(f"{os.path.basename(file)} - Sheet: {sheet_name} dibaca.")
                            awb_list.append(df)
                    else:
                        df = pd.read_excel(file, header=header_row, dtype=str)
                        print(f"Excel dibaca: {os.path.basename(file)}")
                        awb_list.append(df)
                except Exception as e:
                    print(f"Gagal membaca {file}: {e}")
        else:
            raise ValueError("Jenis file tidak didukung: gunakan 'csv' atau 'excel'.")

        combined = []
        for df in awb_list:
            if df.empty:
                print("Data kosong, melewati...")
                continue
            
            # Normalisasi nama kolom
            df.columns = df.columns.str.strip().str.upper()
            awb_col = awb_column.upper()
            status_col = status_column.upper()
            print("Unique STATUS values:", df[status_col].unique()[:50])

            if {awb_col, status_col}.issubset(df.columns):
                df[status_col] = df[status_col].astype(str).str.strip().str.upper()
                df_filtered = df[~df[status_col].isin([s.upper() for s in exclude_statuses])][[awb_col]]
                df_filtered = df_filtered.dropna(subset=[awb_col])
                if not df_filtered.empty:
                    combined.append(df_filtered)
            else:
                pass

        if combined:
            result_df = pd.concat(combined, ignore_index=True)

            if quote_awb:
                result_df[awb_col] = "'" + result_df[awb_col].astype(str)

            result_df.to_csv(output_path, index=False, header=False, encoding="utf-8-sig")
            print(f"Open AWB disimpan di: {output_path}")

            if tracker and task_name:
                tracker.set_preprocess(task_name, True)

            return True
        else:
            print("Tidak ada data AWB yang valid ditemukan.")
            return False

    except Exception as e:
        print(f"Gagal mengekstrak open AWB: {e}")
        return False


def run_open_awb_extraction(open_awb_tasks: list, tracker=None):
    for task in open_awb_tasks:
        task_name = task["desc"]
        print(f"Mengekstrak Open AWB untuk {task_name}...")
        extract_open_awb(
            input_path=task["input_path"],
            output_path=task["output_path"],
            file_type=task["file_type"],
            status_column=task["status_column"],
            awb_column=task["awb_column"],
            exclude_statuses=task.get("exclude_statuses", []),
            header_row=task.get("header_row", 0),
            quote_awb=task.get("quote_awb", False),
            process_all_sheets=task.get("process_all_sheets", False),
            tracker=tracker,
            task_name=task_name
        )
