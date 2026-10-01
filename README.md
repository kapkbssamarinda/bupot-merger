# Web-App Merger Bukti Potong PPh (Multi-Versi: V1 Standar, V2 BPBS, & V3 BPBS)

Aplikasi web untuk mengekstrak, mendeteksi otomatis versi, dan menggabungkan data Bukti Pemotongan/Pemungutan PPh dari file PDF ke dalam format Excel (.xlsx) siap pakai berumus.

## Fitur Utama

- **Deteksi Otomatis Format / Versi Bukti Potong:**
  - **Versi 1 (V1 - BPPU Standar):** Bukti Pemotongan/Pemungutan PPh Unifikasi Berformat Standar DJP.
  - **Versi 2 (V2 - BPBS):** Formulir BPBS PPh Pasal 4 Ayat (2), Pasal 15, Pasal 22, dan Pasal 23 (Format 15 digit angka terpisah spasi).
  - **Versi 3 (V3 - BPBS Baru / NITKU):** Formulir BPBS PPh dengan penambahan field NITKU dan format NPWP Ganda (15 digit lama / 16 digit baru/NITKU).

- **Ekstraksi Otomatis 16 Kolom Terkonsolidasi (Header Excel CAPSLOCK):**
  1. `VERSI` (Identifikasi otomatis V1 Standar / V2 BPBS / V3 BPBS)
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
  12. `TANGGAL DOKUMEN` (Format Short Date: `dd/mm/yyyy`)
  13. `NOMOR DOKUMEN`
  14. `NPWP PEMOTONG` (Identitas Pemotong/Pemungut PPh: V1 16-digit, V2 15-digit, V3 15-digit pertama sebelum `/`)
  15. `NAMA PEMOTONG` (Nama Pemotong: V1 C.3, V2 C.2, V3 C.3)
  16. `TANGGAL PEMOTONGAN` (Format Short Date: `dd/mm/yyyy`)

- **Filter Interaktif & Multi-Upload:**
  - Tombol filter cepat untuk menyaring tampilan: **Semua**, **V1 Standar**, **V2 BPBS**, atau **V3 BPBS**.
  - Area **Drag & Drop** dan tombol **"Pilih Berkas PDF"** untuk memproses banyak file PDF sekaligus.

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
  - Kolom Tanggal Dokumen dan Tanggal Pemotongan diformat **Short Date** (`dd/mm/yyyy`) sebagai objek tanggal Excel asli.
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
