open_awb_tasks = [
    {
            "desc": "Open Danamon",
            "input_path": "D:/RYAN/3. Reports/Danamon/Raw Report.csv",
            "output_path": "c:/Users/DELL/Desktop/ReportApp/data/Open AWB Danamon.csv",
            "file_type": "csv",
            "status_column": "STATUS_POD",
            "awb_column": "AWB",
            "exclude_statuses": ["SUCCESS", "RETURN SHIPPER", "DELIVERED"],
         }
]

new_awb_tasks = []

rt_awb_tasks = []

list_customer_ids = []

apex_config = {
    "upload_endpoint": r"C:\Users\DELL\Desktop\ReportApp\data",
    "download_endpoint": r"C:\Users\DELL\Desktop\ReportApp\data\apex_downloads",
    "archive_folder": r"C:\Users\DELL\Desktop\ReportApp\data\Archive",
}