import os
import shutil

def ensure_folder(path):
    if not os.path.exists(path):
        os.makedirs(path)

def clear_folder_of_csvs(folder):
    for file in os.listdir(folder):
        if file.endswith(".csv"):
            os.remove(os.path.join(folder, file))

def move_csvs_to_folder(source_dir, destination_dir):
    print(f"Mencari file CSV di: {source_dir}")
    for file in os.listdir(source_dir):
        if file.endswith(".csv"):
            src = os.path.join(source_dir, file)
            dst = os.path.join(destination_dir, file)
            print(f"Memindahkan {file} ke {destination_dir}")
            shutil.move(src, dst)
