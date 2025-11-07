# Regresi Linear dan Transformasi

Repositori ini berisi skrip Python untuk melakukan analisis regresi satu variabel terhadap data pada file Excel `Data Analisis Regresi Satu Variabel.xlsx`.

## Menjalankan Analisis

1. Buka terminal dan pindah ke folder tempat repositori ini berada. Jika Anda telah mengunduh atau mengekstrak repositori, gunakan jalur folder aktual di komputer Anda, contohnya:
   ```bash
   cd /Users/nama_pengguna/Documents/Regresi-Linear-dan-Transformasi
   ```
   Sesuaikan jalur di atas dengan lokasi sebenarnya dari folder repositori. Tidak perlu membuat folder baru.

2. Pastikan Python 3 tersedia, kemudian jalankan skrip analisis:
   ```bash
   python onevar_analysis.py
   ```

3. Setelah perintah selesai, hasil analisis dapat ditemukan pada direktori `outputs/`:
   - `outputs/onevar_regression_summary.txt` berisi ringkasan korelasi dan pemilihan model terbaik.
   - Folder `outputs/plots/` menyimpan grafik scatter plot dalam format SVG.

Jika Anda menjalankan skrip dari lokasi lain, Anda dapat memberikan argumen jalur file secara eksplisit:
```bash
python onevar_analysis.py --data "path/ke/Data Analisis Regresi Satu Variabel.xlsx" --output-dir hasil
```
Perintah tersebut akan membaca data dari file yang ditentukan dan menulis seluruh keluaran ke folder `hasil`.

## Dependensi

Skrip hanya menggunakan modul standar Python sehingga tidak memerlukan pemasangan pustaka tambahan.

