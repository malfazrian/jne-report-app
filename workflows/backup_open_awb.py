import os
from core.file_ops import ensure_folder, clear_folder_of_csvs, move_csvs_to_folder

def backup_open_awb_files():
    """Backup file open AWB ke folder Archive dan bersihkan arsip sebelumnya."""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    archive_dir = os.path.join(current_dir, "..", "data", "archive")
    archive_dir = os.path.abspath(archive_dir)  # Normalisasi path

    ensure_folder(archive_dir)
    clear_folder_of_csvs(archive_dir)
    move_csvs_to_folder(current_dir, archive_dir)

    print("Backup selesai.")
