# KlikSurat v4.0

Sistem Manajemen Surat Digital berbasis Google Apps Script yang dirancang khusus untuk efisiensi administrasi persuratan sekolah. Aplikasi ini mengubah Google Sheets menjadi database relasional yang mendukung pembuatan surat otomatis, pengarsipan digital, dan verifikasi dokumen via QR Code.

## 🚀 Fitur Utama

- **Rich Text Editor (TinyMCE 7):** Penulisan redaksi surat yang fleksibel seperti Microsoft Word.
- **Auto-Pagination Engine:** Sistem cerdas yang membagi konten surat ke beberapa halaman secara otomatis (WYSWYG).
- **Verifikasi QR Code:** Setiap surat memiliki Serial ID unik yang dapat diverifikasi keasliannya melalui pemindaian QR.
- **Database Terintegrasi:** Manajemen data Siswa dan Guru dengan fitur impor/ekspor Excel (SheetJS).
- **Sistem Template:** Simpan draf surat yang sering digunakan sebagai master template.
- **Multi-Signature Support:** Mendukung hingga 3 pejabat penandatangan dengan berbagai formasi tata letak.
- **Auto-Numbering:** Penomoran surat dinas otomatis berdasarkan kode klasifikasi dan urutan tahun berjalan.

## 🛠️ Tech Stack

- **Backend:** Google Apps Script (GAS)
- **Database:** Google Sheets
- **Frontend:** HTML5, Tailwind CSS 3.4
- **Libraries:** 
  - TinyMCE 7 (Editor)
  - QRCode.js (Generator QR)
  - SheetJS / XLSX (Data Processing)
  - SweetAlert2 (UI Dialog)
  - html2canvas (Sertifikat Verifikasi)

## 📋 Panduan Instalasi

### 1. Persiapan Database
1. Buat Google Spreadsheet baru.
2. Salin **ID Spreadsheet** Anda (terdapat di URL: `https://docs.google.com/spreadsheets/d/ID_INI/edit`).
3. Buat folder baru di Google Drive untuk menyimpan logo, lalu salin **ID Folder** tersebut.

### 2. Pengaturan Skrip
1. Di Spreadsheet, buka menu **Extensions** > **Apps Script**.
2. Buat file-file berikut di editor Apps Script dan tempelkan kode dari repositori ini:
   - `kode.gs` (dari `kode.js`)
   - `index.html`
   - `modals.html`
   - `sidebar.html`
   - `css.html`
   - `js.html`
   - `verification.html`

### 3. Konfigurasi Keamanan (PENTING)
Agar ID Spreadsheet tidak bocor di repositori publik, jangan tulis ID langsung di kode. Gunakan **Script Properties**:
1. Di editor Apps Script, klik ikon roda gigi (**Project Settings**).
2. Gulir ke bawah ke bagian **Script Properties**.
3. Tambahkan properti berikut:
   - `SPREADSHEET_ID`: (ID Spreadsheet Anda)
   - `FOLDER_LOGO_ID`: (ID Folder Drive Anda)
   - `NPSN_SEKOLAH`: (NPSN Sekolah Anda)

### 4. Deploy Aplikasi
1. Klik tombol **Deploy** > **New Deployment**.
2. Pilih type: **Web App**.
3. Setel akses:
   - *Execute as:* Me (Akun Anda)
   - *Who has access:* Anyone (untuk mendukung verifikasi publik) atau sesuai kebutuhan organisasi.
4. Klik **Deploy** dan salin URL Web App yang muncul.

## 🔐 Keamanan & Kontribusi

### Jangan Memasukkan ke GitHub:
- Jangan mengunggah file `.clasp.json` jika menggunakan CLASP.
- Pastikan `SPREADSHEET_ID` dan kredensial akun default di fungsi `inisialisasiDatabase()` sudah disensor atau dipindahkan ke `Script Properties`.

### Kontribusi
Jika Anda ingin mengembangkan fitur baru:
1. Fork repositori ini.
2. Buat branch fitur baru (`git checkout -b feature/FiturBaru`).
3. Commit perubahan Anda (`git commit -m 'Add some FiturBaru'`).
4. Push ke branch tersebut (`git push origin feature/FiturBaru`).
5. Buat Pull Request.

## 📄 Lisensi

Distribusi terbatas untuk lingkungan pendidikan. Dikembangkan oleh **Muhammad Fahmy**.

---
*© 2026 KlikSurat Engine - Digitalizing Education.*
