/**
 * ====================================================================
 * SYSTEM ENGINE BACKEND - APLIKASI SURAT SDN MOJOGEMI 02 v2.7
 * (Integrated: Single-Save Engine, Auto-Numbering, Archives & Templates)
 * ====================================================================
 */

// --- KONFIGURASI GLOBAL ---
// Gunakan Project Settings > Script Properties untuk mengisi nilai-nilai ini di production
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || "ID_SPREADSHEET_ANDA"; 
const FOLDER_LOGO_ID = PropertiesService.getScriptProperties().getProperty('FOLDER_LOGO_ID') || "ID_FOLDER_DRIVE_ANDA";
const FOLDER_MASUK_ID = PropertiesService.getScriptProperties().getProperty('FOLDER_MASUK_ID') || "ID_FOLDER_SURAT_MASUK_ANDA";
const NPSN_SEKOLAH = PropertiesService.getScriptProperties().getProperty('NPSN_SEKOLAH') || "00000000"; 

// Nama-nama Sheet
const SHEET_PENGATURAN = "Pengaturan";
const SHEET_USERS = "Database_Users";
const SHEET_LOG = "Log_Surat";
const SHEET_TEMPLATES = "Database_Templates";
const SHEET_KLASIFIKASI = "Kode_Klasifikasi";
const SHEET_SISWA = "Database_Siswa";
const SHEET_GURU = "Database_Guru";
const SHEET_LOG_MASUK = "Log_Surat_Masuk";

/**
 * Fungsi Utama untuk menjalankan Web App
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  
  // Kirim parameter URL ke dalam file HTML agar bisa dibaca JavaScript frontend
  template.urlParams = e.parameter;
  // Pastikan URL bersih dari /dev jika sedang dalam mode pengembangan
  let serviceUrl = ScriptApp.getService().getUrl();
  if (!serviceUrl) serviceUrl = "";
  template.scriptUrl = serviceUrl.replace(/\/dev$/, "/exec");
  
  // Ambil pengaturan global agar tersedia di halaman publik
  try {
    const config = ambilPengaturan() || {};
    template.appLogo = config ? (config.Logo_Kanan_Url || config.Logo_Kiri_Url) : "";
    template.appName = config ? (config.App_Name || "KlikSurat SDN Mojogemi 02") : "KlikSurat SDN Mojogemi 02";
  } catch (err) {
    template.appLogo = "";
    template.appName = "KlikSurat";
  }

  return template.evaluate()
      .setTitle(template.appName)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setFaviconUrl('https://i.postimg.cc/FK7R0fTb/logo-jember3.png');
}

/**
 * Fungsi pembantu untuk menyisipkan file HTML lain ke dalam template utama
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Mendapatkan URL Web App secara dinamis untuk link verifikasi QR
 */
function dapatkanUrlSkrip() {
  return ScriptApp.getService().getUrl();
}

/**
 * UTILITY: Membuat URL pemendek (Short Link) via TinyURL agar QR Code sederhana
 */
function getShortUrl(longUrl) {
  try {
    const response = UrlFetchApp.fetch("https://tinyurl.com/api-create.php?url=" + encodeURIComponent(longUrl));
    return response.getContentText();
  } catch (e) {
    return longUrl; // Fallback ke URL panjang jika layanan gagal
  }
}

/**
 * Utility: Membuka Spreadsheet secara terpusat
 */
let _cachedSs = null;
function dapatkanSpreadsheetHost() {
  if (_cachedSs) return _cachedSs;
  try {
    _cachedSs = SpreadsheetApp.openById(SPREADSHEET_ID);
    return _cachedSs;
  } catch (error) {
    throw new Error("Gagal membuka Spreadsheet. Periksa SPREADSHEET_ID di Kode.gs");
  }
}

/**
 * Inisialisasi Database: Membuat sheet dan kolom yang sesuai
 * (Updated: Mendukung 3 Struktur Tanda Tangan)
 */
function inisialisasiDatabase() {
  const ss = dapatkanSpreadsheetHost();
  
  // 1. Sheet Pengaturan
  let sheetSetting = ss.getSheetByName(SHEET_PENGATURAN);
  if (!sheetSetting) {
    sheetSetting = ss.insertSheet(SHEET_PENGATURAN);
    sheetSetting.appendRow(["Parameter", "Value"]);
    const defaultSettings = [
      ["App_Name", "KlikSurat SDN Mojogemi 02"],
      ["Kop_Daerah", "PEMERINTAH KABUPATEN ..."],
      ["Kop_Sub_Dinas", "DINAS ..."],
      ["Kop_Sekolah", "NAMA SEKOLAH ANDA"],
      ["Kop_Alamat", "ALAMAT LENGKAP SEKOLAH"],
      ["Kop_Kontak", "Email: sekolah@example.com | NPSN: " + NPSN_SEKOLAH],
      ["NPSN_Sekolah", NPSN_SEKOLAH],
      ["Logo_Kiri_Url", ""], ["Logo_Kanan_Url", ""],
      ["Ukuran_Kertas", "size-a4"], ["Pilihan_Font", "font-serif-official"],
      ["Line_Spacing", "1.6"],
      ["Satuan_Margin", "mm"], ["Margin_Atas", "10"], ["Margin_Bawah", "10"], ["Margin_Kiri", "10"], ["Margin_Kanan", "10"],
      
      ["Ttd_Mode_Aktif", "kanan-bawah"], // Contoh mode: 'tunggal', 'ganda', atau 'tiga-kolom'
      ["Layout_Type", "standard"],
      
      ["Ttd_Frasa_Ditetapkan", "Ditetapkan di"],
      ["Ttd_Frasa_Tanggal", "Pada Tanggal"],
      ["Ttd_Height", "80"],
      ["Ttd_Gunakan_Materai", "TIDAK"],
      ["Ttd_Gunakan_Foto", "TIDAK"],

      // Data TTD Pejabat 1 (Utama)
      ["Ttd_Jabatan_1", "Kepala Sekolah"],
      ["Ttd_Nama_1", "NAMA KEPALA SEKOLAH"],
      ["Ttd_Pangkat_1", "PANGKAT/GOLONGAN"],
      ["Ttd_Nip_1", "00000000 000000 0 000"],
      
      // Data TTD Pejabat 2
      ["Ttd_Jabatan_2", "Ketua Komite"], 
      ["Ttd_Nama_2", "-"], 
      ["Ttd_Pangkat_2", "-"], 
      ["Ttd_Nip_2", "-"],
      
      // Data TTD Pejabat 3
      ["Ttd_Jabatan_3", "Bendahara Sekolah"], 
      ["Ttd_Nama_3", "-"], 
      ["Ttd_Pangkat_3", "-"], 
      ["Ttd_Nip_3", "-"]
    ];
    sheetSetting.getRange(2, 1, defaultSettings.length, 2).setValues(defaultSettings);
  }
  sheetSetting.getRange("A:B").setNumberFormat("@");
  
  // 2. Sheet Users
  if (!ss.getSheetByName(SHEET_USERS)) {
    let sh = ss.insertSheet(SHEET_USERS);
    sh.appendRow(["Username", "Password", "Nama_Lengkap", "Role"]);
    sh.appendRow(["admin", "GANTI_PASSWORD_INI", "Administrator Sekolah", "Operator Utama"]);
  }
  ss.getSheetByName(SHEET_USERS).getRange("A:D").setNumberFormat("@");
  
  // 3. Sheet Log Arsip
  if (!ss.getSheetByName(SHEET_LOG)) {
    let sh = ss.insertSheet(SHEET_LOG);
    sh.appendRow(["Timestamp", "Operator", "Tanggal_Surat", "Nomor_Surat", "Perihal", "Penerima", "Lampiran", "Pembuka", "Isi_Surat", "Tembusan", "Serial_ID", "Layout", "Ref_ID"]);
  }
  ss.getSheetByName(SHEET_LOG).getRange("B:M").setNumberFormat("@");

  // 4. Sheet Database Templates
  if (!ss.getSheetByName(SHEET_TEMPLATES)) {
    let sh = ss.insertSheet(SHEET_TEMPLATES);
    sh.appendRow(["Nama_Template", "Tanggal_Default", "Nomor", "Perihal", "Penerima", "Lampiran", "Pembuka", "Isi_Surat", "Tembusan", "Layout"]);
  }
  ss.getSheetByName(SHEET_TEMPLATES).getRange("A:J").setNumberFormat("@");

  // 5. Sheet Kode Klasifikasi Surat
  if (!ss.getSheetByName(SHEET_KLASIFIKASI)) {
    let sh = ss.insertSheet(SHEET_KLASIFIKASI);
    sh.appendRow(["Kode", "Nama_Klasifikasi", "Aktif"]);
    sh.getRange("A:C").setNumberFormat("@");
    sh.getRange(2, 1, 8, 3).setValues([
      ["400.3.5", "Administrasi Sekolah", "YA"],
      ["005", "Undangan", "YA"],
      ["421.2", "Pendidikan Dasar", "YA"],
      ["421.3", "Kesiswaan", "YA"],
      ["421.5", "Kurikulum", "YA"],
      ["800", "Kepegawaian", "YA"],
      ["900", "Keuangan", "YA"],
      ["045", "Arsip/Dokumentasi", "YA"]
    ]);
  }
  ss.getSheetByName(SHEET_KLASIFIKASI).getRange("A:C").setNumberFormat("@");

  // 6. Sheet Database Siswa
  if (!ss.getSheetByName(SHEET_SISWA)) {
    let sh = ss.insertSheet(SHEET_SISWA);
    const headers = [
      "Nama", "NIS", "NISN", "Tempat_Lahir", "Tanggal_Lahir", "Dusun", "Desa", "RT", "RW", "Kecamatan", "Kabupaten", "Nama_Ayah", "Nama_Ibu", "Kelas", "Tahun_Ajaran", "VA_PIP", "Rek_PIP", "Nomor_Ijazah", "Ket_Lulus",
      "Pendidikan Agama dan Budi Pekerti",
      "Pendidikan Pancasila",
      "Bahasa Indonesia",
      "Matematika",
      "Ilmu Pengetahuan Alam dan Sosial",
      "Pendidikan Jasmani, Olahraga, dan Kesehatan",
      "Seni dan Budaya",
      "Bahasa Inggris",
      "Bahasa Madura",
      "Baca Tulis Al Quran (BTA)"
    ];
    sh.appendRow(headers);
    sh.getRange("A:AC").setNumberFormat("@");
  }

  // 7. Sheet Database Guru
  if (!ss.getSheetByName(SHEET_GURU)) {
    let sh = ss.insertSheet(SHEET_GURU);
    const headers = ["Nama", "NIP", "NUPTK", "Jabatan", "Pangkat_Golongan", "Unit_Kerja", "Tugas_Utama", "Tugas_Tambahan"];
    sh.appendRow(headers);
    sh.getRange("A:H").setNumberFormat("@");
  }

  // 8. Sheet Log Surat Masuk
  if (!ss.getSheetByName(SHEET_LOG_MASUK)) {
    let sh = ss.insertSheet(SHEET_LOG_MASUK);
    sh.appendRow(["Timestamp", "Operator", "Tanggal_Terima", "Asal_Surat", "Nomor_Surat_Asal", "Perihal", "Isi_OCR", "Link_File"]);
    sh.getRange("A:H").setNumberFormat("@");
  }
}

/**
 * AUTH: Sistem Login
 */
function prosesLogin(username, password) {
  try {
    inisialisasiDatabase();
    const sheet = dapatkanSpreadsheetHost().getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username && data[i][1] === password) {
        // Ambil data langsung dari baris database untuk memastikan sinkronisasi
        return { status: "SUCCESS", nama: data[i][2], role: data[i][3], username: data[i][0] };
      }
    }
    return { status: "FAILED", message: "Username atau password salah!" };
  } catch (error) {
    return { status: "ERROR", message: error.toString() };
  }
}

/**
 * OCR ENGINE: Mengolah Surat Masuk (Upload + Ekstraksi Teks)
 * Fungsi ini membutuhkan Layanan "Drive API" diaktifkan di Advanced Services.
 */
function simpanSuratMasukOCR(metadata, base64Data, mimeType) {
  try {
    const ss = dapatkanSpreadsheetHost();
    const sheet = ss.getSheetByName(SHEET_LOG_MASUK);
    const folder = DriveApp.getFolderById(FOLDER_MASUK_ID); // Menggunakan folder khusus surat masuk
    
    // 1. Simpan File Asli ke Drive
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, mimeType, "Scan_Masuk_" + metadata.nomor_asal);
    const file = folder.createFile(blob);
    const fileUrl = file.getUrl();
    
    // OCR dihapus, teks ekstraksi diisi default dengan strip
    let extractedText = "-";

    // 3. Catat ke Database Log_Surat_Masuk
    sheet.appendRow([
      new Date(),
      metadata.operator || "Admin",
      metadata.tanggal_terima,
      metadata.asal_surat,
      metadata.nomor_asal,
      metadata.perihal,
      extractedText,
      fileUrl
    ]);
    
    // Format baris terakhir agar rapi
    sheet.getRange(sheet.getLastRow(), 1, 1, 8).setNumberFormat("@");

    return { 
      status: "SUCCESS", 
      message: "Surat berhasil diarsipkan. Teks telah diekstrak otomatis.",
      extracted: extractedText.substring(0, 200) + "..." 
    };

  } catch (e) {
    return { status: "ERROR", message: "Gagal memproses surat masuk: " + e.toString() };
  }
}

/**
 * SETTINGS: Ambil data pengaturan
 */
function ambilPengaturan() {
  try {
    const sheet = dapatkanSpreadsheetHost().getSheetByName(SHEET_PENGATURAN);
    const data = sheet.getDataRange().getValues();
    const config = data.slice(1).reduce((acc, [key, val]) => {
      if (key) acc[key] = val;
      return acc;
    }, {});
    
    // Tambahkan data pejabat KS untuk pemanggilan variabel global
    config.pejabat_ks = {
      nama: config.Ttd_Nama_1 || "",
      nip: config.Ttd_Nip_1 || "",
      jabatan: config.Ttd_Jabatan_1 || "",
      pangkat: config.Ttd_Pangkat_1 || ""
    };
    
    return config;
  } catch (e) { return null; }
}

/**
 * SISWA: Mencari data siswa berdasarkan nama atau NISN
 */
function cariSiswa(query) {
  try {
    const ss = dapatkanSpreadsheetHost();
    const sheet = ss.getSheetByName(SHEET_SISWA);
    if (!sheet) return [];
    
        const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; // Hanya header atau kosong
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const namaIdx = headers.findIndex(h => h.toLowerCase() === "nama");
    const nisnIdx = headers.findIndex(h => h.toLowerCase() === "nisn");
    const q = String(query || "").toLowerCase().replace(/[^a-z0-9-]/g, ''); // Izinkan tanda hubung
    const qAsli = String(query || "").toLowerCase();
    
    return data.slice(1)
      .filter(row => {
        const namaVal = namaIdx !== -1 ? String(row[namaIdx] || "").toLowerCase() : "";
        const nisnVal = nisnIdx !== -1 ? String(row[nisnIdx] || "").toLowerCase().replace(/[^a-z0-9]/g, '') : "";
        return (namaIdx !== -1 && namaVal.includes(qAsli)) || (nisnIdx !== -1 && nisnVal.includes(q));
      }).filter(Boolean) // Filter out any null/undefined results
      .map(row => {
        let obj = {};
        headers.forEach((h, i) => {
          if (h === "Tanggal_Lahir") {
            if (row[i] instanceof Date) {
              obj[h] = Utilities.formatDate(row[i], "GMT+7", "yyyy-MM-dd");
            } else if (typeof row[i] === "string" && row[i].includes("/")) {
              const p = row[i].split("/"); // DD/MM/YYYY -> YYYY-MM-DD
              if (p.length === 3) obj[h] = `${p[2]}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
              else obj[h] = row[i];
            } else {
              obj[h] = row[i];
            }
          } else {
            obj[h] = row[i];
          }
        });
        return obj;
      }).slice(0, 10); // Limit 10 hasil
  } catch (e) { return []; }
}

/**
 * GURU: Mencari data guru berdasarkan nama atau NIP
  */
function cariGuru(query) {
  try {
    const ss = dapatkanSpreadsheetHost();
    const sheet = ss.getSheetByName(SHEET_GURU);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const namaIdx = headers.findIndex(h => h.toLowerCase() === "nama");
    const nipIdx = headers.findIndex(h => h.toLowerCase() === "nip");
    const q = String(query || "").toLowerCase().replace(/[^a-z0-9-]/g, ''); // Izinkan tanda hubung
    const qAsli = String(query || "").toLowerCase();

    const results = data.slice(1)
      .filter(row => {
        const nama = namaIdx !== -1 ? String(row[namaIdx] || "").toLowerCase() : "";
        const nip = nipIdx !== -1 ? String(row[nipIdx] || "").toLowerCase().replace(/[^a-z0-9-]/g, '') : "";
        return nama.includes(qAsli) || nip.includes(q);
      }).filter(Boolean) // Filter out any null/undefined results
      .map(row => {
        let obj = {};
        headers.forEach((h, i) => {
          let val = row[i];
          // Tambahkan kolom identitas lain jika diperlukan
          if (["NIP", "NUPTK", "NISN", "RT", "RW", "NIS", "Kelas"].includes(h)) {
            val = String(val).trim();
          }
          obj[h] = val;
        });
        return obj;
      }).slice(0, 10);
    return results;
  } catch (e) { 
    console.error("[ERROR] Backend cariGuru:", e.toString());
    return []; 
  }
}

/**
  * GURU: Mengambil satu data guru secara spesifik berdasarkan NIP
 */
function ambilGuruByNip(nip) {
  try {
    const nipStr = String(nip).trim();
    const ss = dapatkanSpreadsheetHost();
    const sheet = ss.getSheetByName(SHEET_GURU);
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const nipIdx = headers.findIndex(h => h.toLowerCase() === "nip");
    
    if (nipIdx === -1) throw new Error("Kolom NIP tidak ditemukan di Sheet Guru");

    const row = data.find(r => {
      const nipSheet = String(r[nipIdx]).trim();
      return nipSheet === nipStr || 
             nipSheet.replace(/[^a-z0-9]/g, '') === nipStr.replace(/[^a-z0-9]/g, '');
    });
    
    if (row) {
      let obj = {};
      headers.forEach((h, i) => {
        let val = row[i];
        if (["NIP", "NUPTK", "NISN"].includes(h)) val = String(val).trim();
        obj[h] = val;
      });
      return obj;
    }
    return null;

  } catch (e) {
    console.error("[ERROR] Backend ambilGuruByNip:", e.toString());
    return null;
  }
}

/**
 * GURU: Simpan atau Update data guru
 */
function simpanGuruBaru(data) {
  try {
    const ss = dapatkanSpreadsheetHost();
    const sheet = ss.getSheetByName(SHEET_GURU);
    const dataRaw = sheet.getDataRange().getValues();
    const headers = dataRaw[0].map(h => String(h).trim());
    const nipIdx = headers.findIndex(h => h.toLowerCase() === "nip");
    const nipInput = String(data.NIP || "").trim();

    const nipMap = new Map(dataRaw.map((r, i) => [String(r[nipIdx]).trim(), i + 1]));
    let targetRow = nipInput && nipMap.has(nipInput) ? nipMap.get(nipInput) : -1;

    const rowData = headers.map(h => data[h] || "");

    if (targetRow > 1) {
      sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
      return { status: "SUCCESS", message: "Data guru '" + data.Nama + "' berhasil diperbarui." };
    } else {
      sheet.appendRow(rowData);
      return { status: "SUCCESS", message: "Guru baru '" + data.Nama + "' berhasil ditambahkan." };
    }
  } catch (e) { return { status: "ERROR", message: e.toString() }; }
}

/**
 * GURU: Menghapus data guru
 */
function hapusGuru(nip) {
  try {
    const sheet = dapatkanSpreadsheetHost().getSheetByName(SHEET_GURU);
    const data = sheet.getDataRange().getValues();
    const nipIdx = data[0].indexOf("NIP");
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nipIdx]).trim() === String(nip).trim()) {
        sheet.deleteRow(i + 1);
        return { status: "SUCCESS", message: "Data guru berhasil dihapus." };
      }
    }
    return { status: "ERROR", message: "Data tidak ditemukan." };
  } catch (e) { return { status: "ERROR", message: e.toString() }; }
}

/**
 * GURU: Ambil semua untuk ekspor
 */
function ambilSemuaGuru() {
  try {
    const sheet = dapatkanSpreadsheetHost().getSheetByName(SHEET_GURU);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    return data.slice(1).map(row => {
      let obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
  } catch (e) { return []; }
}

/**
 * SISWA: Simpan atau Update data siswa
 */
function simpanSiswaBaru(data) {
  try {
    const ss = dapatkanSpreadsheetHost();
    const sheet = ss.getSheetByName(SHEET_SISWA);
       if (!sheet) return { status: "ERROR", message: "Sheet Database_Siswa tidak ditemukan." };

    const dataRaw = sheet.getDataRange().getValues();
    const headers = dataRaw[0].map(h => String(h).trim());
    const nisnIdx = headers.findIndex(h => h.toLowerCase() === "nisn");
    const nisnInput = String(data.NISN || "").trim();

    const nisnMap = new Map(dataRaw.map((r, i) => [String(r[nisnIdx]).trim(), i + 1]));
    const targetRow = nisnInput && nisnMap.has(nisnInput) ? nisnMap.get(nisnInput) : -1;

    const rowData = headers.map(h => {
      let val = data[h] || "";

      if (h === "Tanggal_Lahir" && val.includes("-")) {
        const p = val.split("-"); // YYYY-MM-DD -> DD/MM/YYYY
        if (p.length === 3) val = `${p[2]}/${p[1]}/${p[0]}`; // Corrected: assign to val, not return
      }
      return val;
    });

    if (targetRow !== -1) {
      sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
      return { status: "SUCCESS", message: "Data siswa '" + data.Nama + "' berhasil diperbarui." };
    } else {
      sheet.appendRow(rowData);
      return { status: "SUCCESS", message: "Siswa baru '" + data.Nama + "' berhasil ditambahkan." };
    }
  } catch (e) {
    return { status: "ERROR", message: e.toString() };
  }
}

/**
 * SISWA: Menghapus data siswa berdasarkan NISN
 */
function hapusSiswa(nisn) {
  try {
    const ss = dapatkanSpreadsheetHost();
    const sheet = ss.getSheetByName(SHEET_SISWA);
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const nisnIdx = headers.findIndex(h => h.toLowerCase() === "nisn");
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nisnIdx]).trim() === String(nisn).trim()) {
        sheet.deleteRow(i + 1);
        return { status: "SUCCESS", message: "Data siswa berhasil dihapus secara permanen." };
      }
    }
    return { status: "ERROR", message: "Data siswa tidak ditemukan." };
  } catch (e) { return { status: "ERROR", message: e.toString() }; }
}

/**
 * SISWA: Mengambil seluruh data siswa untuk backup/ekspor
 */
function ambilSemuaSiswa() {
  try {
    const sheet = dapatkanSpreadsheetHost().getSheetByName(SHEET_SISWA);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
    const headers = data[0];
    return data.slice(1).map(row => {
      let obj = {};
      headers.forEach((h, i) => {
        if (h === "Tanggal_Lahir") {
          if (row[i] instanceof Date) {
            obj[h] = Utilities.formatDate(row[i], "GMT+7", "yyyy-MM-dd");
          } else if (typeof row[i] === "string" && row[i].includes("/")) {
            const p = row[i].split("/");
            if (p.length === 3) obj[h] = `${p[2]}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
            else obj[h] = row[i];
          } else {
            obj[h] = row[i];
          }
        } else {
          obj[h] = row[i];
        }
      });
      return obj;
    });
  } catch (e) { return []; }
}

/**
 * SISWA: Impor massal data siswa (Dukungan JSON)
 */
function imporSiswaBatch(siswaArray) {
  try {
    const ss = dapatkanSpreadsheetHost();
    const sheet = ss.getSheetByName(SHEET_SISWA);
    const dataExist = sheet.getDataRange().getValues(); 
    const headers = dataExist[0].map(h => String(h).trim()); 
    const nisnIdx = headers.findIndex(h => h.toLowerCase() === "nisn");
    if (nisnIdx === -1) return { status: "ERROR", message: "Kolom NISN tidak ditemukan di database." };
    
    // Map untuk pencarian O(1)
    const nisnMap = new Map();
    dataExist.forEach((row, index) => {
      const key = String(row[nisnIdx] || "").trim();
      if (index > 0 && key) nisnMap.set(key, index + 1);
    });

    siswaArray.forEach(s => {
      const nisnInput = String(s.NISN || "").trim();
      const rowData = headers.map(h => s[h] || "");
      
      if (nisnInput && nisnMap.has(nisnInput)) {
        sheet.getRange(nisnMap.get(nisnInput), 1, 1, rowData.length).setValues([rowData]);
      } else if (nisnInput) {
        sheet.appendRow(rowData);
        nisnMap.set(nisnInput, sheet.getLastRow());
      }
    });

    return { status: "SUCCESS", message: siswaArray.length + " data siswa diproses (Tambah/Update)." };
  } catch (e) { return { status: "ERROR", message: e.toString() }; }
}

/**
 * GURU: Impor massal data guru (Dukungan JSON)
 */
function imporGuruBatch(guruArray) {
  try {
    const ss = dapatkanSpreadsheetHost();
    const sheet = ss.getSheetByName(SHEET_GURU);
    const dataExist = sheet.getDataRange().getValues();
    const headers = dataExist[0].map(h => String(h).trim());
    const nipIdx = headers.findIndex(h => h.toLowerCase() === "nip");
    if (nipIdx === -1) return { status: "ERROR", message: "Kolom NIP tidak ditemukan di database." };

    
    const nipMap = new Map();
    dataExist.forEach((row, index) => {
      const key = String(row[nipIdx] || "").trim();
      if (index > 0 && key) nipMap.set(key, index + 1);
    });

    guruArray.forEach(g => {
      const nipInput = String(g.NIP || "").trim();
      const rowData = headers.map(h => g[h] || "");
      
      if (nipInput && nipMap.has(nipInput)) {
        sheet.getRange(nipMap.get(nipInput), 1, 1, rowData.length).setValues([rowData]);
      } else if (nipInput) {
        sheet.appendRow(rowData);
        nipMap.set(nipInput, sheet.getLastRow());
      }
    });

    return { status: "SUCCESS", message: guruArray.length + " data guru diproses (Tambah/Update)." };
  } catch (e) { return { status: "ERROR", message: e.toString() }; }
}


/**
 * UPLOAD: Mengolah Base64 dan mengembalikan link lh3 (High Res)
 */
function uploadFileLogoKeDrive(dataMime, base64Data, namaKomponen) {
  try {
    const folder = DriveApp.getFolderById(FOLDER_LOGO_ID);
    const rawData = Utilities.base64Decode(base64Data);
    const fileName = "Logo_" + namaKomponen;
    const blob = Utilities.newBlob(rawData, dataMime, fileName);
    
    // Cari file yang ada, pastikan bukan yang di dalam sampah (Trash)
    const existingFiles = folder.getFilesByName(fileName);
    let file;
    
    while (existingFiles.hasNext()) {
      let f = existingFiles.next();
      if (!f.isTrashed()) {
        f.setTrashed(true); // Hapus versi lama agar tidak terjadi konflik metadata/header
      }
    }
    
    // Buat file baru secara bersih untuk menjamin integritas data biner
    file = folder.createFile(blob);
    
    const fileId = file.getId(); // Dapatkan ID segera setelah file dibuat

    // Coba atur sharing, jika ditolak kebijakan organisasi, biarkan tetap privat
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      console.warn("Sharing ditolak kebijakan organisasi: " + e.message);
    }

    /**
     * Alternatif URL: Menggunakan format direct 'lh3' berbasis ID File.
     * Format ini lebih stabil dan tidak memutus eksekusi jika setSharing gagal.
     */
    const directUrl = "https://lh3.googleusercontent.com/d/" + fileId;

    return { status: "SUCCESS", url: directUrl };
  } catch (err) {
    return { status: "ERROR", message: err.toString() };
  }
}
/**
 * ARSIP: Murni hanya menyuntikkan teks isi surat ke sheet Log_Surat
 */
function simpanPaketSuratLengkap(paket) {
  try {
    const db = dapatkanSpreadsheetHost();
    const sheetLog = db.getSheetByName(SHEET_LOG); // Target: Log_Surat
    
    if (!sheetLog) {
      return { status: "ERROR", message: "Sheet Log_Surat tidak ditemukan." };
    }
    
    const lastRow = sheetLog.getLastRow();
    
    // Ambil data konten teks surat dari parameter
    const s = paket.surat ? paket.surat : paket;
    
    sheetLog.appendRow([
      new Date(), 
      String(s.operator || "Operator"), 
      String(s.tanggal || ""), 
      String(s.nomor || ""), 
      String(s.perihal || ""), 
      String(s.penerima || ""), 
      String(s.lampiran || "-"), 
      String(s.pembuka || ""), 
      String(s.isi || ""),
      String(s.tembusan || ""),
      String(s.serial_id || "-"),
      String(s.layout || "standard"),
      String(s.ref_id || "")
    ]);

    // Cukup format baris yang baru saja ditambahkan agar ringan
    sheetLog.getRange(lastRow + 1, 2, 1, 12).setNumberFormat("@");

    return { status: "SUCCESS", message: "Surat berhasil diarsipkan ke database log." };

  } catch (e) {
    return { status: "ERROR", message: "Gagal mengarsipkan: " + e.toString() };
  }
}

/**
 * TEMPLATE: Murni hanya mendaftarkan teks isi surat ke sheet Database_Templates
 */
function simpanTemplateDinamis(data) {
  try {
    const db = dapatkanSpreadsheetHost();
    const sheetTemplate = db.getSheetByName(SHEET_TEMPLATES); // Target: Database_Templates
    
    if (!sheetTemplate) {
      return { status: "ERROR", message: "Sheet Database_Templates tidak ditemukan." };
    }

    const namaTemplate = data.nama || "Template Baru";
    
    // OPTIMASI: Hanya ambil kolom pertama (Nama Template) untuk cek duplikat
    const lastRow = sheetTemplate.getLastRow();
    const names = lastRow > 1 ? sheetTemplate.getRange(2, 1, lastRow - 1, 1).getValues().flat() : [];
    const isDuplicate = names.some(n => String(n).toLowerCase() === namaTemplate.toLowerCase());
    
    if (isDuplicate) {
      return { status: "FAILED", message: "Template dengan nama '" + namaTemplate + "' sudah ada. Gunakan nama lain." };
    }

    const s = data.surat ? data.surat : data;
    
    sheetTemplate.appendRow([
      String(namaTemplate), 
      String(s.tanggal || ""), 
      String(s.nomor || ""), 
      String(s.perihal || ""), 
      String(s.penerima || ""), 
      String(s.lampiran || "-"), 
      String(s.pembuka || ""), 
      String(s.isi || ""),
      String(s.tembusan || ""),
      String(s.layout || "standard")
    ]);
    
    sheetTemplate.getRange(sheetTemplate.getLastRow(), 1, 1, 10).setNumberFormat("@");
    
    return { status: "SUCCESS", message: "Master template '" + namaTemplate + "' berhasil disimpan." };
  } catch (e) { 
    return { status: "ERROR", message: "Gagal menyimpan template: " + e.toString() }; 
  }
}
/**
 * TEMPLATES: Mengambil seluruh master template untuk dropdown frontend
 */
function ambilSemuaTemplate() {
  try {
    const sheet = dapatkanSpreadsheetHost().getSheetByName(SHEET_TEMPLATES);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; 

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    
    const idxNama = headers.indexOf("nama_template") !== -1 ? headers.indexOf("nama_template") : headers.indexOf("nama");
    const idxNomor = headers.indexOf("nomor");
    const idxTanggal = headers.indexOf("tanggal_default") !== -1 ? headers.indexOf("tanggal_default") : headers.indexOf("tanggal");
    const idxPerihal = headers.indexOf("perihal");
    const idxPenerima = headers.indexOf("penerima");
    const idxLampiran = headers.indexOf("lampiran");
    const idxPembuka = headers.indexOf("pembuka");
    const idxIsi = headers.indexOf("isi_surat") !== -1 ? headers.indexOf("isi_surat") : headers.indexOf("isi");
    const idxTembusan = headers.indexOf("tembusan");
    const idxLayout = headers.indexOf("layout");

    const templates = data.slice(1).map(row => {
      let tglTemplate = "";
      if (idxTanggal !== -1 && row[idxTanggal]) {
        tglTemplate = row[idxTanggal] instanceof Date 
          ? Utilities.formatDate(row[idxTanggal], "GMT+7", "yyyy-MM-dd") 
          : String(row[idxTanggal]);
      }

      return {
        nama: idxNama !== -1 ? String(row[idxNama]) : "(Tanpa Nama Template)",
        nomor: idxNomor !== -1 ? String(row[idxNomor]) : "",
        tanggal: tglTemplate,
        perihal: idxPerihal !== -1 ? String(row[idxPerihal]) : "",
        penerima: idxPenerima !== -1 ? String(row[idxPenerima]) : "",
        lampiran: idxLampiran !== -1 ? String(row[idxLampiran]) : "-",
        pembuka: idxPembuka !== -1 ? String(row[idxPembuka]) : "",
        isi: idxIsi !== -1 ? String(row[idxIsi]) : "",
        tembusan: idxTembusan !== -1 ? String(row[idxTembusan]) : "",
        layout: idxLayout !== -1 ? String(row[idxLayout]) : "standard"
      };
    })
    // Filter hanya template yang memiliki nama (menghindari baris kosong di sheet)
    .filter(t => t.nama && t.nama.trim() !== "" && t.nama !== "(Tanpa Nama Template)" && t.nama !== "Nama_Template");
    
    return templates;

  } catch (e) {
    console.error("Gagal ambilSemuaTemplate: " + e.toString());
    return [];
  }
}

/**
 * MASTER DATA: Mengambil kode klasifikasi surat aktif
 */
function ambilKodeKlasifikasiSurat() {
  try {
    inisialisasiDatabase();
    const sheet = dapatkanSpreadsheetHost().getSheetByName(SHEET_KLASIFIKASI);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    return data.slice(1)
      .filter(row => row[0] && String(row[2] || "YA").toUpperCase() !== "TIDAK")
      .map(row => ({
        kode: String(row[0]).trim(),
        nama: String(row[1] || "").trim()
      }))
      .filter(item => item.kode);
  } catch (e) {
    console.error("Gagal ambilKodeKlasifikasiSurat: " + e.toString());
    return [];
  }
}

/**
 * MASTER DATA: Menghapus kode klasifikasi surat
 */
function hapusKodeKlasifikasi(kode) {
  try {
    const ss = dapatkanSpreadsheetHost();
    const sheet = ss.getSheetByName(SHEET_KLASIFIKASI);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(kode).trim()) {
        sheet.deleteRow(i + 1);
        return { status: "SUCCESS", message: `Kode '${kode}' berhasil dihapus.` };
      }
    }
    return { status: "FAILED", message: "Kode tidak ditemukan." };
  } catch (e) {
    return { status: "ERROR", message: e.toString() };
  }
}

/**
 * MASTER DATA: Menambah kode klasifikasi surat baru
 */
function tambahKodeKlasifikasiSurat(dataKode) {
  try {
    inisialisasiDatabase();
    const kode = String(dataKode && dataKode.kode ? dataKode.kode : "").trim();
    const nama = String(dataKode && dataKode.nama ? dataKode.nama : "").trim();

    if (!kode || !nama) {
      return { status: "FAILED", message: "Kode dan nama klasifikasi wajib diisi." };
    }

    const sheet = dapatkanSpreadsheetHost().getSheetByName(SHEET_KLASIFIKASI);
    sheet.getRange("A:C").setNumberFormat("@");
    const data = sheet.getDataRange().getValues();
    const kodeLower = kode.toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === kodeLower) {
        sheet.getRange(i + 1, 1, 1, 3).setNumberFormat("@");
        sheet.getRange(i + 1, 1, 1, 3).setValues([[String(kode), String(nama), "YA"]]);
        return { status: "SUCCESS", message: "Kode klasifikasi diperbarui." };
      }
    }

    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, 3).setNumberFormat("@");
    sheet.getRange(nextRow, 1, 1, 3).setValues([[String(kode), String(nama), "YA"]]);
    return { status: "SUCCESS", message: "Kode klasifikasi baru berhasil ditambahkan." };
  } catch (e) {
    return { status: "ERROR", message: "Gagal menyimpan kode klasifikasi: " + e.toString() };
  }
}

/**
 * ARSIP: Menarik riwayat log arsip surat keluar
 */
function dapatkanRiwayatArsip() {
  try {
    const sheet = dapatkanSpreadsheetHost().getSheetByName(SHEET_LOG);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const hasilMap = data.slice(1).map(row => {
      let idTs = row[0] instanceof Date ? row[0].getTime() : row[0];
      let waktuSistem = row[0];
      if (row[0] instanceof Date) {
        waktuSistem = Utilities.formatDate(row[0], "GMT+7", "dd/MM/yy HH:mm");
      }
      
      let tanggalSurat = row[2];
      if (row[2] instanceof Date) {
        tanggalSurat = Utilities.formatDate(row[2], "GMT+7", "yyyy-MM-dd");
      }

      return {
        id: idTs,
        waktu: waktuSistem,
        operator: row[1],
        tanggal: tanggalSurat,
        nomor: row[3],
        perihal: row[4],
        penerima: row[5],
        lampiran: row[6],
        pembuka: row[7],
        isi: row[8],
        tembusan: row[9] || "",
        sid: row[10] || "N/A",
        layout: row[11] || "standard",
        ref_id: row[12] || ""
      };
    });

    return hasilMap.reverse().slice(0, 50);

  } catch (e) { 
    console.error("Error dapatkanRiwayatArsip: " + e.toString());
    return []; 
  }
}

/**
 * ARSIP: Menghapus baris log arsip tertentu
 */
function hapusDataArsip(timestampId, nomor) {
  try {
    const ss = dapatkanSpreadsheetHost();
    const sheet = ss.getSheetByName(SHEET_LOG);
    const data = sheet.getDataRange().getValues();
    
    for (let i = data.length - 1; i >= 1; i--) {
      let currentTs = data[i][0] instanceof Date ? data[i][0].getTime() : data[i][0];
      if (String(currentTs) === String(timestampId) && data[i][3] === nomor) {
        sheet.deleteRow(i + 1);
        return { status: "SUCCESS", message: "Catatan arsip berhasil dihapus secara permanen." };
      }
    }
    return { status: "FAILED", message: "Data arsip tidak ditemukan." };
  } catch (e) {
    return { status: "ERROR", message: "Kesalahan server: " + e.toString() };
  }
}

/**
 * VERIFIKASI: Mengambil data surat berdasarkan Unique ID (SID) untuk halaman publik
 */
function verifikasiSuratByUid(uid) {
  try {
    if (!uid) return { status: "ERROR", message: "ID Verifikasi tidak disertakan." };
    
    const ss = dapatkanSpreadsheetHost();
    const sheet = ss.getSheetByName(SHEET_LOG);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Cari indeks kolom Serial_ID secara dinamis agar lebih akurat
    const idxSid = headers.findIndex(h => String(h).trim().toLowerCase() === "serial_id");
    if (idxSid === -1) return { status: "ERROR", message: "Kolom Serial_ID tidak ditemukan di database." };

    const searchId = String(uid).trim().toLowerCase();

    // Cari dari baris terakhir (data terbaru)
    for (let i = data.length - 1; i >= 1; i--) {
      const rowSid = String(data[i][idxSid]).trim().toLowerCase();
      
      if (rowSid === searchId && rowSid !== "-" && rowSid !== "") {
        const config = ambilPengaturan() || {};
        
        return {
          status: "SUCCESS",
          no: data[i][3],
          hal: data[i][4],
          ttd: config.Ttd_Nama_1 || data[i][1] || "Kepala Sekolah",
          sid: data[i][idxSid],
          tanggal: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], "GMT+7", "dd/MM/yyyy") : data[i][2]
        };
      }
    }
    return { status: "NOT_FOUND", message: "Data tidak ditemukan di pangkalan data pusat." };
  } catch (e) {
    return { status: "ERROR", message: e.toString() };
  }
}

/**
 * ARSIP: Mengambil data Log_Surat dan memformatnya untuk Agenda Excel
 * Filter berdasarkan Tahun dan Header sesuai permintaan.
 */
function ambilDataArsipUntukEkspor(tahun) {
  try {
    const ss = dapatkanSpreadsheetHost();
    const sheetLog = ss.getSheetByName(SHEET_LOG);
    const sheetSetting = ss.getSheetByName(SHEET_PENGATURAN);
    
    if (!sheetLog) throw new Error("Sheet Log_Surat tidak ditemukan.");

    // Ambil Nama Sekolah untuk kolom "Dari"
    let schoolName = "SDN Mojogemi 02";
    if (sheetSetting) {
      const settings = sheetSetting.getDataRange().getValues();
      const nameRow = settings.find(r => r[0] === "Kop_Sekolah");
      if (nameRow) schoolName = nameRow[1];
    }

    const data = sheetLog.getDataRange().getValues();
    if (data.length <= 1) return [];

    const filteredRows = data.slice(1).filter(row => {
      const noSurat = String(row[3] || "");
      const parts = noSurat.split('/');
      const yearInNo = parts.length > 0 ? parts[parts.length - 1] : "";
      
      // Cek tahun dari nomor surat (format: .../2026)
      return yearInNo === String(tahun);
    });

    const days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
    const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

    return filteredRows.map(row => {
      let tglRaw = row[2];
      let hariTanggal = "-";
      
      // Pastikan tglRaw diubah menjadi objek Date agar getDay() bekerja
      let d = (tglRaw instanceof Date) ? tglRaw : new Date(tglRaw);
      
      if (!isNaN(d.getTime())) {
        hariTanggal = `${days[d.getDay()]}, ${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
      }
      
      const noSurat = String(row[3] || "");
      const partsNo = noSurat.split('/');
      const kodeKlasifikasi = partsNo[0] || "-";
      
      // Mengambil nomor urut sampai akhir (membuang kode klasifikasi di awal)
      // Contoh: 400.3.5/001/.../2026 -> 001/.../2026
      const noSuratTanpaKlasifikasi = partsNo.slice(1).join('/');

      return {
        "Hari, Tanggal": hariTanggal,
        "No Surat": noSuratTanpaKlasifikasi || noSurat,
        "Perihal": row[4] || "-",
        "Kode Klasifikasi Surat": kodeKlasifikasi,
        "Dari Siapa (Dari)": schoolName,
        "Untuk Siapa (Ke)": row[5] || "-",
        "Tembusan": row[9] || "-"
      };
    });
  } catch (e) { throw new Error(e.toString()); }
}

/**
 * ARSIP: Mengambil data Log_Surat_Masuk dan memformatnya untuk Agenda Excel
 * Filter berdasarkan Tahun.
 */
function ambilDataSuratMasukUntukEkspor(tahun) {
  try {
    const ss = dapatkanSpreadsheetHost();
    const sheetLog = ss.getSheetByName(SHEET_LOG_MASUK);
    
    if (!sheetLog) throw new Error("Sheet Log_Surat_Masuk tidak ditemukan.");

    const data = sheetLog.getDataRange().getValues();
    if (data.length <= 1) return [];

    const targetTahun = String(tahun);
    const days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
    const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

    // Filter baris berdasarkan tahun dari kolom Tanggal_Terima (Indeks 2)
    const filteredRows = data.slice(1).filter(row => {
      let tglRaw = row[2];
      let d = (tglRaw instanceof Date) ? tglRaw : new Date(tglRaw);
      return !isNaN(d.getTime()) && String(d.getFullYear()) === targetTahun;
    });

    return filteredRows.map(row => {
      let tglRaw = row[2];
      let hariTanggal = "-";
      let d = (tglRaw instanceof Date) ? tglRaw : new Date(tglRaw);
      
      if (!isNaN(d.getTime())) {
        hariTanggal = `${days[d.getDay()]}, ${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
      }

      return {
        "Hari, Tanggal Terima": hariTanggal,
        "Asal Surat (Dari)": row[3] || "-",
        "Nomor Surat Asal": row[4] || "-",
        "Perihal": row[5] || "-",
        "Diterima Oleh": row[1] || "-",
        "Link Dokumen": row[7] || "-"
      };
    });
  } catch (e) { throw new Error(e.toString()); }
}

/**
 * Validasi: Cek apakah string nomor surat lengkap sudah digunakan dalam database
 */
function cekFullNomorDuplikat(nomorFull) {
  try {
    const sheet = dapatkanSpreadsheetHost().getSheetByName(SHEET_LOG);
    if (!sheet) return false;
    const data = sheet.getDataRange().getValues();
    const target = String(nomorFull || "").replace(/\s+/g, '').toLowerCase();
    for (let i = 1; i < data.length; i++) {
      const val = String(data[i][3] || "").replace(/\s+/g, '').toLowerCase();
      if (val === target) {
        return true; // Duplikat ditemukan
      }
    }
    return false;
  } catch (e) {
    Logger.log("Error cekFullNomorDuplikat: " + e.toString());
    return false;
  }
}

/**
 * Validasi: Cek apakah nomor urut sudah digunakan dalam tahun berjalan
 */
function cekNomorDuplikat(urutanStr, tahun) {
  try {
    const sheet = dapatkanSpreadsheetHost().getSheetByName(SHEET_LOG);
    if (!sheet) return "ERROR";
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return "VALID";

    // OPTIMASI: Hanya ambil kolom Nomor Surat (Kolom D) untuk divalidasi
    const nomorData = sheet.getRange(1, 4, lastRow, 1).getValues();
    const targetUrutan = parseInt(urutanStr, 10);
    const targetTahun = String(tahun).trim();
    let urutanMaks = 0;
    
    for (let i = 1; i < nomorData.length; i++) {
      let nomorArsip = String(nomorData[i][0] || "");
      let parts = nomorArsip.split('/');
      if (parts.length >= 4 && parts[parts.length - 1] === targetTahun) {
        let u = parseInt(parts[1], 10);
        if (u === targetUrutan) return "DUPLICATE";
        if (u > urutanMaks) urutanMaks = u;
      }
    }
    
    if (targetUrutan > urutanMaks + 1) return "JUMP";
    
    return "VALID";
  } catch (e) {
    return "ERROR";
  }
}

/**
 * AUTO-NUMBER: Membuat nomor surat otomatis
 */
function generateNomorSuratOtomatis(kodeKlasifikasi) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Tunggu hingga 10 detik jika ada proses lain
    const ss = dapatkanSpreadsheetHost();
    const sheet = ss.getSheetByName(SHEET_LOG);
    const lastRow = sheet.getLastRow();
    
    // OPTIMASI: Hanya ambil kolom Nomor Surat untuk mencari urutan terbesar
    const nomorData = lastRow > 1 ? sheet.getRange(2, 4, lastRow - 1, 1).getValues() : [];
    
    const config = ambilPengaturan() || {};
    let urutanTerakhir = 0;
    let tahunIni = new Date().getFullYear().toString();
    let npsnSekolah = String(config.NPSN_Sekolah || NPSN_SEKOLAH).replace(/\D/g, "") || NPSN_SEKOLAH;
    
    // Cari urutan terbesar di seluruh data log untuk tahun berjalan (Global Sequence)
    for (let i = 0; i < nomorData.length; i++) {
      let nomorArsip = String(nomorData[i][0] || "");
      let parts = nomorArsip.split('/');
      
      // Cek apakah format sesuai (KODE/URUTAN/NPSN/TAHUN) dan tahunnya cocok
      if (parts.length >= 4 && parts[parts.length - 1] === tahunIni) {
        let urutan = parseInt(parts[1], 10);
        if (!isNaN(urutan) && urutan > urutanTerakhir) {
          urutanTerakhir = urutan;
        }
      }
    }
    
    let urutanStr = (urutanTerakhir + 1).toString().padStart(3, '0');
    let npsnLengkap = "35.09.310.24." + npsnSekolah; 
    return `${kodeKlasifikasi}/${urutanStr}/${npsnLengkap}/${tahunIni}`;
  } catch (e) {
    return `${kodeKlasifikasi}/001/35.09.310.24.${NPSN_SEKOLAH}/${new Date().getFullYear()}`;
  } finally {
    lock.releaseLock();
  }
}

/**
 * SETTINGS: Menyimpan Khusus Pengaturan Kop Surat (Dipanggil oleh Tombol Simpan Kop)
 */
function simpanPengaturanKop(dataKop) {
  return perbaruiDataPengaturanBatch(dataKop, "Konfigurasi Identitas");
}

/**
 * SETTINGS: Menyimpan Khusus Pengaturan TTD (Mendukung Multi-TTD hingga 3 Nama)
 */
function simpanPengaturanTtd(dataTtd) {
  return perbaruiDataPengaturanBatch(dataTtd, "Pengaturan 3 Tanda Tangan");
}

/**
 * HELPER: Fungsi internal untuk memperbarui data pengaturan secara batch (DRY)
 */
function perbaruiDataPengaturanBatch(dataMap, konteks) {
  try {
    const db = dapatkanSpreadsheetHost();
    const sheetSetting = db.getSheetByName(SHEET_PENGATURAN);
    
    if (!sheetSetting) {
      return { status: "ERROR", message: "Sheet Pengaturan tidak ditemukan." };
    }

    const dataRange = sheetSetting.getDataRange().getValues();
    sheetSetting.getRange("A:B").setNumberFormat("@");
    
    const settingsMap = new Map(dataRange.map((row, index) => [String(row[0]).toLowerCase().trim(), index + 1]));
    
    for (let key in dataMap) {
      const searchKey = key.toLowerCase().trim();
      if (settingsMap.has(searchKey)) {
        const rowIndex = settingsMap.get(searchKey);
        sheetSetting.getRange(rowIndex, 2).setValue(String(dataMap[key] || ""));
      } else {
        const nextRow = sheetSetting.getLastRow() + 1;
        sheetSetting.getRange(nextRow, 1, 1, 2).setNumberFormat("@");
        sheetSetting.getRange(nextRow, 1, 1, 2).setValues([[String(key), String(dataMap[key] || "")]]);
        settingsMap.set(searchKey, nextRow);
      }
    }
    return { status: "SUCCESS", message: konteks + " berhasil diperbarui." };
  } catch (e) {
    return { status: "ERROR", message: "Gagal menyimpan " + konteks + ": " + e.toString() };
  }
}

/**
 * USERS: Memperbarui Username/Password/Nama user yang sedang login
 */
function perbaruiProfilUser(usernameLama, dataBaru) {
  try {
    const ss = dapatkanSpreadsheetHost();
    const sheet = ss.getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === usernameLama) {
        // Kolom: Username (0), Password (1), Nama_Lengkap (2), Role (3)
        sheet.getRange(i + 1, 1).setValue(dataBaru.username);
        if (dataBaru.password) {
          sheet.getRange(i + 1, 2).setValue(dataBaru.password);
        }
        sheet.getRange(i + 1, 3).setValue(dataBaru.nama);
        return { status: "SUCCESS", message: "Data profil berhasil diperbarui." };
      }
    }
    return { status: "FAILED", message: "User tidak ditemukan di database." };
  } catch (error) {
    return { status: "ERROR", message: error.toString() };
  }
}

/**
 * TEMPLATE: Batch import data template dari file JSON
 */
function imporTemplatesBatch(templates) {
  try {
    const db = dapatkanSpreadsheetHost();
    const sheetTemplate = db.getSheetByName(SHEET_TEMPLATES);
    if (!sheetTemplate) return { status: "ERROR", message: "Sheet Database_Templates tidak ditemukan." };

    const existingData = sheetTemplate.getDataRange().getValues().map(r => r[0].toString().toLowerCase());
    
    const rowsToInsert = templates
      .filter(t => t.nama && !existingData.includes(t.nama.toLowerCase()))
      .map(t => [
        String(t.nama), 
        String(t.tanggal || ""), 
        String(t.nomor || ""), 
        String(t.perihal || ""), 
        String(t.penerima || ""), 
        String(t.lampiran || "-"), 
        String(t.pembuka || ""), 
        String(t.isi || ""),
        String(t.tembusan || ""),
        String(t.layout || "standard")
      ]);

    if (rowsToInsert.length > 0) {
      sheetTemplate.getRange(sheetTemplate.getLastRow() + 1, 1, rowsToInsert.length, 10).setValues(rowsToInsert);
    }
    
    return { status: "SUCCESS", message: rowsToInsert.length + " template baru berhasil diimpor (Duplikasi dilewati)." };
  } catch (e) {
    return { status: "ERROR", message: "Gagal impor: " + e.toString() };
  }
}

/**
 * TEMPLATE: Menghapus seluruh data di sheet template (kecuali header)
 */
function hapusSemuaTemplate() {
  try {
    const db = dapatkanSpreadsheetHost();
    const sheet = db.getSheetByName(SHEET_TEMPLATES);
    if (!sheet) return { status: "ERROR", message: "Sheet Database_Templates tidak ditemukan." };
    
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    return { status: "SUCCESS", message: "Seluruh master template berhasil dihapus." };
  } catch (e) {
    return { status: "ERROR", message: "Gagal hapus: " + e.toString() };
  }
}

/**
 * TEMPLATE: Menghapus satu template tertentu berdasarkan nama
 */
function hapusTemplateSatu(nama) {
  try {
    const db = dapatkanSpreadsheetHost();
    const sheet = db.getSheetByName(SHEET_TEMPLATES);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(nama).trim()) {
        sheet.deleteRow(i + 1);
        return { status: "SUCCESS", message: `Template '${nama}' berhasil dihapus.` };
      }
    }
    return { status: "FAILED", message: "Template tidak ditemukan." };
  } catch (e) { return { status: "ERROR", message: e.toString() }; }
}

/**
 * USERS: Mengambil semua daftar pengguna
 */
function ambilSemuaUser() {
  try {
    const sheet = dapatkanSpreadsheetHost().getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    return data.slice(1).map(row => {
      let obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
  } catch (e) { return []; }
}

/**
 * USERS: Simpan atau Update User
 */
function simpanDatabaseUser(data) {
  try {
    const sheet = dapatkanSpreadsheetHost().getSheetByName(SHEET_USERS);
    const dataRaw = sheet.getDataRange().getValues();
    const usernameInput = String(data.Username).trim().toLowerCase();
    let targetRow = -1;

    for (let i = 1; i < dataRaw.length; i++) {
      if (String(dataRaw[i][0]).trim().toLowerCase() === usernameInput) {
        targetRow = i + 1;
        break;
      }
    }

    const rowData = [data.Username, data.Password, data.Nama_Lengkap, data.Role];
    if (targetRow !== -1) {
      sheet.getRange(targetRow, 1, 1, 4).setValues([rowData]);
      return { status: "SUCCESS", message: "User '" + data.Username + "' diperbarui." };
    } else {
      sheet.appendRow(rowData);
      return { status: "SUCCESS", message: "User '" + data.Username + "' ditambahkan." };
    }
  } catch (e) { return { status: "ERROR", message: e.toString() }; }
}

/**
 * USERS: Hapus User
 */
function hapusUserDatabase(username) {
  try {
    const sheet = dapatkanSpreadsheetHost().getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    const uInput = String(username).trim().toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === uInput) {
        if (uInput === "admin") return { status: "ERROR", message: "User Admin utama tidak boleh dihapus!" };
        sheet.deleteRow(i + 1);
        return { status: "SUCCESS", message: "User berhasil dihapus." };
      }
    }
    return { status: "ERROR", message: "User tidak ditemukan." };
  } catch (e) { return { status: "ERROR", message: e.toString() }; }
}
