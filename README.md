# ReportApp

Aplikasi ini digunakan untuk mengelola, mengekstrak, menggabungkan, mengunggah, dan mengunduh laporan AWB dari berbagai sumber secara otomatis.

## Fitur
- Ekstraksi data Open AWB, New AWB, dan RT AWB dari file CSV/Excel.
- Penggabungan file CSV hasil ekstraksi.
- Backup file sebelum proses baru.
- Upload dan download data ke/dari sistem Apex secara otomatis.
- Tracking status setiap task (preprocess, request, download).

## Requirement

- **Python 3.10+** (disarankan Python 3.11 atau 3.13)
- **Google Chrome** (untuk Selenium)
- **ChromeDriver** yang sesuai versi Chrome
- Paket Python:
  - pandas
  - selenium
  - openpyxl

### Cara Install Requirement

```sh
pip install pandas selenium openpyxl
```

> **Pastikan ChromeDriver sudah ada di PATH atau satu folder dengan Python Anda.**

## Cara Menjalankan

1. **Edit file task**  
   Pastikan file di folder `tasks/ryan_tasks.py` sudah sesuai dengan kebutuhan Anda (path file, customer_id, dsb).

2. **Jalankan aplikasi**  
   Buka terminal di folder `ReportApp`, lalu jalankan:

   ```sh
   python main.py
   ```

   atau jika menggunakan Windows dan Python 3.13:

   ```sh
   C:/Users/DELL/AppData/Local/Programs/Python/Python313/python.exe main.py
   ```

3. **Ikuti proses di terminal**  
   Semua proses dan status akan tampil di terminal, termasuk summary tracker di akhir.

## Struktur Folder

```
ReportApp/
│
├─ main.py
├─ tasks/
│   └─ ryan_tasks.py
├─ workflows/
│   ├─ tracker.py
│   ├─ file_ops.py
│   └─ extractors/
│       ├─ open_awb_extractor.py
│       ├─ new_awb_extractor.py
│       ├─ awb_rt_extractor.py
│       └─ apex_data_extractor.py
└─ data/
    ├─ (hasil ekstrak dan download)
```

## Catatan

- Pastikan semua path di file task sudah benar dan file sumber tersedia.
- Jika ingin menambah/mengubah task, edit file `tasks/ryan_tasks.py`.
- Untuk troubleshooting Selenium/ChromeDriver, cek error di terminal.

---

**Selamat