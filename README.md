# Web-App Merger Bukti Potong PPh (Dual-Versi: V1 Standar & V2 BPBS)

Aplikasi web untuk mengekstrak, mendeteksi otomatis versi, dan menggabungkan data Bukti Pemotongan/Pemungutan PPh dari file PDF ke dalam format Excel (.xlsx) siap pakai berumus.

## Fitur Utama

- **Deteksi Otomatis Format / Versi Bukti Potong:**
  - **Versi 1 (V1 - BPPU Standar):** Bukti Pemotongan/Pemungutan PPh Unifikasi Berformat Standar DJP.
  - **Versi 2 (V2 - BPBS):** Formulir BPBS PPh Pasal 4 Ayat (2), Pasal 15, Pasal 22, dan Pasal 23.

- **Ekstraksi Otomatis 16 Kolom Terkonsolidasi (Header Excel CAPSLOCK):**
  1. `VERSI` (Identifikasi otomatis V1 Standar / V2 BPBS)
  2. `NOMOR BUPOT` (Nomor Bukti Potong)
  3. `MASA PAJAK`
  4. `SIFAT PAJAK PENGHASILAN` (Final / Tidak Final)
  5. `STATUS` (Normal / Pembetulan)
  6. `JENIS PPH` (Pasal 22, Pasal 23, Pasal 4(2), Pasal 15, dll.)
  7. `KODE OBJEK PAJAK`
  8. `OBJEK PAJAK` (Deskripsi / Keterangan Kode Objek)
  9. `DPP` (Dasar Pengenaan Pajak)
  10. `TARIF` (%)
  11. `PPH DIPOTONG` (Pajak Penghasilan Terpotong)
  12. `TANGGAL DOKUMEN`
  13. `NOMOR DOKUMEN`
  14. `NPWP PEMOTONG` (Identitas Pemotong/Pemungut PPh)
  15. `NAMA PEMOTONG` (Nama Badan/Wajib Pajak Pemotong)
  16. `TANGGAL PEMOTONGAN`

- **Filter Interaktif & Multi-Upload:**
  - Tombol filter cepat untuk menyaring tampilan: **Semua**, **V1 Standar**, atau **V2 BPBS**.
  - Area **Drag & Drop** dan tombol **"Tambah Dokumen PDF"** untuk memproses banyak file PDF sekaligus.

- **Antarmuka Anti-AI Slop & Tema Terang (Light Theme):**
  - Desain editorial finansial profesional (bersih, kontras tinggi ≥ 7:1).
  - Tipografi angka tabular (*tabular numbers*) untuk akurasi data finansial.
  - Kartu ringkasan agregat otomatis (Total Dokumen, Rincian V1 & V2, Total DPP, Total PPh).
  - Pencarian interaktif dan pratinjau tabel sebelum diunduh.

- **Ekspor Excel Siap Pakai (.xlsx):**
  - Kolom `VERSI` memperjelas jenis bukti potong tiap baris.
  - Header berpenampilan profesional dengan latar hijau petroleum dan teks putih tebal.
  - Kolom DPP & PPh diformat angka ribuan (`#,##0`) agar bisa langsung dirumus.
  - Kolom NPWP diformat Teks (`@`) sehingga `00...` tidak terpotong.
  - Baris `TOTAL` otomatis di bagian bawah dengan rumus Excel `=SUM(...)`.
  - Lebar kolom otomatis disesuaikan (*auto-fit*).

## Cara Menjalankan Aplikasi

1. Buka terminal (PowerShell atau Command Prompt) di direktori ini.
2. Jalankan perintah:
   ```bash
   python app.py
   ```
3. Buka browser dan akses alamat:
   ```
   http://localhost:5000
   ```
4. Tarik-lepas atau pilih berkas PDF bukti potong (baik V1 maupun V2) yang ingin diekstrak.
5. Periksa pratinjau tabel, filter versi bila diperlukan, lalu klik **"Unduh Excel (.xlsx)"** untuk menyimpan hasil rekapan.
