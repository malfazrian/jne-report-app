import csv
import os

class TaskTableTracker:
    def __init__(self, task_list):
        # task_list: list of dict, minimal ada 'desc' atau 'customer_id'
        self.rows = []
        for task in task_list:
            row = {
                "task": task.get("desc") or task.get("nama_customer"),
                "preprocess": False,
                "request": False,
                "download": False,
                "result_path": None,
                "praprocess": False
            }
            self.rows.append(row)

    def set_preprocess(self, task_name, value=True):
        for row in self.rows:
            if str(row["task"]) == str(task_name):
                row["preprocess"] = value

    def set_request(self, task_name, value=True):
        for row in self.rows:
            if str(row["task"]) == str(task_name):
                row["request"] = value

    def set_download(self, task_name, value=True):
        for row in self.rows:
            if str(row["task"]) == str(task_name):
                row["download"] = value

    def set_path(self, task_name, value=None):
        for row in self.rows:
            if str(row["task"]) == str(task_name):
                row["result_path"] = value

    def set_praprocess(self, task_name, value=True):
        for row in self.rows:
            if str(row["task"]) == str(task_name):
                row["praprocess"] = value

    def summary(self, output_folder="data/tracker", filename="tracker_summary.csv"):
        print("\n=== TASK TABLE TRACKER ===")
        print(f"{'Task':<20} {'Preprocess':<10} {'Request':<10} {'Download':<10} {'Praprocess':<10}")
        for row in self.rows:
            print(f"{row['task']:<20} {str(row['preprocess']):<10} {str(row['request']):<10} {str(row['download']):<10} {str(row['praprocess']):<10}")
        
        os.makedirs(output_folder, exist_ok=True)
        output_path = os.path.join(output_folder, filename)
        with open(output_path, mode="w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Task", "Preprocess", "Request", "Download", "Result Path", "Praprocess"])
            for row in self.rows:
                writer.writerow([
                    row["task"],
                    row["preprocess"],
                    row["request"],
                    row["download"],
                    row.get("result_path", ""),
                    row["praprocess"]
                ])
        print(f"Tracker summary disimpan ke: {output_path}")

# Contoh inisialisasi:
# from tasks.ryan_tasks import open_awb_tasks
# tracker = TaskTableTracker(open_awb_tasks)