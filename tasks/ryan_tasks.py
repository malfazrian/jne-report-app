open_awb_tasks = [
    {
            "desc": "Open Danamon",
            "input_path": "D:/RYAN/3. Reports/Danamon/Raw Report.csv",
            "output_path": "c:/Users/DELL/Desktop/ReportApp/data/Open AWB Danamon.csv",
            "file_type": "csv",
            "status_column": "STATUS_POD",
            "awb_column": "AWB",
            "exclude_statuses": ["SUCCESS", "RETURN SHIPPER", "DELIVERED"],
        },
        {
            "desc": "Open Generali",
            "input_path": "D:/RYAN/3. Reports/Generali",
            "output_path": "c:/Users/DELL/Desktop/ReportApp/data/Open AWB Generali.csv",
            "file_type": "excel",
            "status_column": "DELIVERY STATUS",
            "awb_column": "AWB",
            "exclude_statuses": ["SUCCESS", "RETURN SHIPPER"],
            "header_row": 2
        },
        {
            "desc": "Open Kalog",
            "input_path": "D:/RYAN/3. Reports/Kalog",
            "output_path": "c:/Users/DELL/Desktop/ReportApp/data/Open AWB Kalog.csv",
            "file_type": "excel",
            "status_column": "STATUS_POD",
            "awb_column": "AWB",
            "exclude_statuses": ["SUCCESS", "RETURN SHIPPER", "DELIVERED"]
        },
        {
            "desc": "Open LSIN",
            "input_path": r"D:\RYAN\3. Reports\LSIN",
            "output_path": r"c:/Users/DELL/Desktop/ReportApp/data/Open AWB LSIN.csv",
            "file_type": "excel",
            "status_column": "STATUS_POD",
            "awb_column": "AWB",
            "exclude_statuses": ["SUCCESS", "RETURN SHIPPER", "DELIVERED"]
        },
        {
            "desc": "Open Smartfren",
            "input_path": r"D:\RYAN\3. Reports\Smartfren",
            "output_path": r"c:/Users/DELL/Desktop/ReportApp/data/Open AWB Smartfren.csv",
            "file_type": "excel",
            "status_column": "STATUS_POD",
            "awb_column": "AWB",
            "exclude_statuses": ["SUCCESS", "RETURN SHIPPER", "DELIVERED"],
            "quote_awb": True,
            "process_all_sheets": True
        },
]

new_awb_tasks = [
    {
        "desc": "New Smartfren",
        "base_folder": r"\\192.168.9.74\f\ALL REPORT\01. JANUARI 2024\ALL REPORT GABUNGAN",
        "archive_file": r"c:/Users/DELL/Desktop/ReportApp/data/Archive/Open AWB Smartfren.csv",
        "output_file": r"c:/Users/DELL/Desktop/ReportApp/data/New AWB Smartfren.csv"
    }
]

rt_awb_tasks = [
    {
        "desc": "RT Danamon",
        "input_path": "D:/RYAN/3. Reports/Danamon/Raw Report.csv",
        "output_path": "c:/Users/DELL/Desktop/ReportApp/data/RT AWB Danamon.csv"
    }
]

list_customer_ids = [
    {
        "customer_id": "81635700",
        "nama_customer": "New Danamon 00",
        "include_today": True,
    },
    {
        "customer_id": "81635701",
        "nama_customer": "New Danamon 01",
        "include_today": True,
    },
    {
        "customer_id": "81728700",
        "nama_customer": "New Kalog",
        "include_today": False,
    },
    {
        "customer_id": "81901200",
        "nama_customer": "New LSIN",
        "include_today": False,
    },
]

apex_config = {
    "upload_endpoint": r"C:\Users\DELL\Desktop\ReportApp\data",
    "download_endpoint": r"C:\Users\DELL\Desktop\ReportApp\data\apex_downloads",
    "archive_folder": r"C:\Users\DELL\Desktop\ReportApp\data\Archive",
}