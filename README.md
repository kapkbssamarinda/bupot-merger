# Web-App Merger Bukti Potong PPh Unifikasi

Aplikasi web untuk mengekstrak dan menggabungkan data Bukti Pemotongan/Pemungutan PPh Unifikasi format standar DJP dari file PDF ke dalam format Excel (.xlsx).

## Fitur Utama

- **Ekstraksi Otomatis 15 Kolom Standar (Header Excel CAPSLOCK):**
  1. `NOMOR`
  2. `MASA PAJAK`
  3. `SIFAT PAJAK PENGHASILAN`
  4. `STATUS`
  5. `JENIS PPH`
  6. `KODE OBJEK PAJAK`
  7. `OBJEK PAJAK`
  8. `DPP`
  9. `TARIF`
  10. `PAJAK PENGHASILAN`
  11. `TANGGAL DOKUMEN`
  12. `NOMOR DOKUMEN`
  13. `NPWP`
  14. `NAMA PEMOTONG`
  15. `TANGGAL PEMOTONGAN`

- **Mekanisme Input PDF:**
  - Area **Drag & Drop** dan tombol **"Pilih File PDF"** untuk memproses satu atau banyak berkas PDF bukti potong secara fleksibel kapan saja.

- **Antarmuka Anti-AI Slop & Tema Terang (Light Theme):**
  - Desain editorial finansial profesional (bersih, kontras tinggi ≥ 7:1).
  - Tipografi angka tabular (*tabular numbers*) untuk akurasi data finansial.
  - Kartu ringkasan agregat otomatis (Total Dokumen, Total DPP, Total PPh).
  - Pencarian interaktif dan pratinjau tabel sebelum diunduh.

- **Ekspor Excel Siap Pakai (.xlsx):**
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
4. Tarik-lepas atau pilih berkas PDF bukti potong unifikasi yang ingin diekstrak.
5. Periksa pratinjau tabel dan ringkasan, lalu klik **"Unduh Excel (.xlsx)"** untuk menyimpan hasil rekapan.
