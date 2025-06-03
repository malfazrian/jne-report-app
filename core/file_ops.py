import os
import shutil

def ensure_folder(path: str):
    """Pastikan folder ada, jika tidak maka buat."""
    if not os.path.exists(path):
        os.makedirs(path)
        print(f"📁 Folder dibuat: {path}")

def clear_folder_of_csvs(folder: str):
    """Hapus semua file .csv dalam folder."""
    for file in os.listdir(folder):
        if file.lower().endswith(".csv"):
            try:
                os.remove(os.path.join(folder, file))
                print(f"🗑️ Dihapus dari {folder}: {file}")
            except Exception as e:
                print(f"❌ Gagal menghapus {file}: {e}")

def move_csvs_to_folder(src_folder: str, dst_folder: str):
    """Pindahkan semua file .csv dari folder sumber ke folder tujuan."""
    for file in os.listdir(src_folder):
        if file.lower().endswith(".csv"):
            src = os.path.join(src_folder, file)
            dst = os.path.join(dst_folder, file)
            try:
                shutil.move(src, dst)
                print(f"📦 Dipindah ke {dst_folder}: {file}")
            except Exception as e:
                print(f"❌ Gagal memindahkan {file}: {e}")
