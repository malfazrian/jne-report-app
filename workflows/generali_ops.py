import os
import xlwings as xw

def process_generali_reports():
    folder_path = r"D:\RYAN\3. Reports\Generali"
    vlookup_source = r"'D:\RYAN\2. Queries\3. Query Report Generali.xlsx'!Data_Input[[AWB]:[RECEIVED/REASON]]"

    lookup_mappings = {
        "DELIVERY STATUS": 26,
        "DATE": 13,
        "NAME": 35,
        "STATUS RECEIVER": 34
    }

    column_names = list(lookup_mappings.keys())

    for file in os.listdir(folder_path):
        if file.endswith(".xlsx") and not file.startswith("~$"):
            file_path = os.path.join(folder_path, file)
            print(f"Memproses: {file_path}")

            app = xw.App(visible=False)
            try:
                wb = app.books.open(file_path)
                sheet = wb.sheets[0]

                header_row = 3
                max_col = 100  # Asumsikan kolom tidak lebih dari 100
                headers = sheet.range((header_row, 1), (header_row, max_col)).value

                if not headers or not isinstance(headers, list):
                    print(f"Header tidak ditemukan atau bukan list di {file_path}, skip.")
                    wb.close()
                    continue

                headers = [str(h).strip() if h else "" for h in headers]
                col_positions = {name: headers.index(name) + 1 for name in column_names if name in headers}
                awb_col_index = headers.index("AWB") + 1 if "AWB" in headers else None

                if not col_positions or not awb_col_index:
                    print(f"Header tidak lengkap di {file_path}, skip.")
                    wb.close()
                    continue

                # Bersihkan header dari None dan spasi
                headers = [str(h).strip() if h else "" for h in headers]

                col_positions = {name: headers.index(name) + 1 for name in column_names if name in headers}
                awb_col_index = headers.index("AWB") + 1 if "AWB" in headers else None

                if not col_positions or not awb_col_index:
                    print(f"Header tidak lengkap di {file_path}, skip.")
                    wb.close()
                    continue

                last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row

                # Tambahkan formula
                for row in range(4, last_row + 1):
                    awb_cell = f"${sheet.cells(row, awb_col_index).address.split('$')[1]}{row}"
                    for col_name, vlookup_col_index in lookup_mappings.items():
                        col_num = col_positions[col_name]
                        cell = sheet.cells(row, col_num)
                        formula = f'=VLOOKUP({awb_cell},{vlookup_source},{vlookup_col_index},FALSE)'
                        cell.value = formula

                wb.app.calculate()

                # Replace formula dengan nilai hasil hitung
                for row in range(4, last_row + 1):
                    for col_name in column_names:
                        col_num = col_positions[col_name]
                        cell = sheet.cells(row, col_num)
                        val = cell.value
                        cell.value = val

                wb.save()
                wb.close()
                print(f"Sukses update dan simpan nilai di: {file_path}")
            except Exception as e:
                print(f"Error pada file {file_path}: {e}")
            finally:
                app.quit()
