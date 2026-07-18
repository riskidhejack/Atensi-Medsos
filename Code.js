/**
 * E-REPORT PUBLIKASI KANIM TANJUNG UBAN (VERSION 2.6 - ULTRA STABLE)
 * Database: Google Sheets (Arsip_Publikasi)
 * AI Engine: Gemini 2.5 Flash API + Secure Local Fallback Compiler
 */

const SPREADSHEET_ID = "1X6zygH4jx9tuXfCMDH0zBirmOMC9rUwhUjzlxwXl3Rc";
const SHEET_NAME = "Arsip_Publikasi";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('E-Report Publikasi Kanim Tanjung Uban')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function dapatkanSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "ID_KONTEN", "TANGGAL_RILIS", "NAMA_KEGIATAN", 
      "LINK_INSTAGRAM", "LINK_FACEBOOK", "LINK_X", 
      "LINK_TIKTOK", "LINK_YT_SHORTS", "LINK_YOUTUBE", 
      "LINK_INSTANSI", "LINK_MEDIA", "WAKTU_INPUT"
    ]);
  }
  return sheet;
}

function simpanKonten(data) {
  try {
    const sheet = dapatkanSheet();
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss");
    const idKonten = "ID-" + timestamp;
    const waktuInput = Utilities.formatDate(new Date(), "GMT+7", "HH:mm:ss");

    sheet.appendRow([
      idKonten,
      data.tanggal,
      data.kegiatan,
      data.instagram || "",
      data.facebook || "",
      data.x || "",
      data.tiktok || "",
      data.yt_shorts || "",
      data.youtube || "",
      data.instansi || "",
      data.media || "",
      waktuInput
    ]);
    return { success: true, message: "Data publikasi berhasil diarsipkan ke database!" };
  } catch (error) {
    return { success: false, message: "Gagal menyimpan: " + error.toString() };
  }
}

function hapusKontenSpesifik(idKonten) {
  try {
    const sheet = dapatkanSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf("ID_KONTEN");

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === idKonten) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Kegiatan berhasil dihapus secara permanen!" };
      }
    }
    return { success: false, message: "Data ID Konten tidak ditemukan." };
  } catch (error) {
    return { success: false, message: "Gagal menghapus: " + error.toString() };
  }
}

function ambilRiwayatKategori() {
  try {
    const sheet = dapatkanSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const rows = data.slice(1);

    let kelompokTanggal = {};

    rows.forEach(row => {
      let tglRaw = row[headers.indexOf("TANGGAL_RILIS")];
      if (!tglRaw) return;
      
      let tglStr = tglRaw;
      if (tglRaw instanceof Date) {
        tglStr = Utilities.formatDate(tglRaw, "GMT+7", "yyyy-MM-dd");
      } else {
        tglStr = tglRaw.toString().trim();
      }

      if (!kelompokTanggal[tglStr]) {
        kelompokTanggal[tglStr] = {
          TANGGAL_RILIS: tglStr,
          DAFTAR_KEGIATAN: []
        };
      }

      kelompokTanggal[tglStr].DAFTAR_KEGIATAN.push({
        id: row[headers.indexOf("ID_KONTEN")] || "ID-MANUAL",
        nama: row[headers.indexOf("NAMA_KEGIATAN")] || "Tanpa Judul"
      });
    });

    return Object.values(kelompokTanggal).sort((a, b) => {
      return new Date(b.TANGGAL_RILIS) - new Date(a.TANGGAL_RILIS);
    });
  } catch (error) {
    return [];
  }
}

function generateLaporanGemini(tanggalTerpilih) {
  try {
    const sheet = dapatkanSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return "Database kosong.";

    const headers = data[0];
    const rows = data.slice(1);

    const kegiatanHariIni = rows.filter(row => {
      let tglRow = row[headers.indexOf("TANGGAL_RILIS")];
      if (!tglRow) return false;
      if (tglRow instanceof Date) {
        tglRow = Utilities.formatDate(tglRow, "GMT+7", "yyyy-MM-dd");
      } else {
        tglRow = tglRow.toString().trim();
      }
      return tglRow === tanggalTerpilih;
    });

    if (kegiatanHariIni.length === 0) return "Tidak ditemukan data publikasi pada tanggal: " + tanggalTerpilih;

    // Urutkan berdasarkan waktu input secara aman
    kegiatanHariIni.sort((a, b) => {
      let wktA = (a[headers.indexOf("WAKTU_INPUT")] || "").toString();
      let wktB = (b[headers.indexOf("WAKTU_INPUT")] || "").toString();
      return wktA.localeCompare(wktB);
    });

    let tanggalFormal = tanggalTerpilih;
    try {
      const opsiTgl = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
      tanggalFormal = new Date(tanggalTerpilih).toLocaleDateString('id-ID', opsiTgl);
    } catch(e) {}

    // =========================================================================
    // RANCANGAN FORMAT INTERNAL SESUAI CONTOH DIREKTORAT JENDERAL IMIGRASI
    // =========================================================================
    let textReportBackup = "Assalamu’alaikum Wr. Wb.\n\n" +
                           "Yth.\n" +
                           "1. Direktur Jenderal Imigrasi;\n" +
                           "2. Sekretaris Direktorat Jenderal Imigrasi;\n" +
                           "3. Para Direktur di lingkungan Direktorat Jenderal Imigrasi;\n" +
                           "4. Kepala Kantor Wilayah Direktorat Jenderal Imigrasi Kepulauan Riau;\n\n" +
                           "Dari: Kepala Kantor Imigrasi Kelas II TPI Tanjung Uban\n\n" +
                           "Bersama ini dengan hormat melaporkan terkait Publikasi Kantor Imigrasi Kelas II TPI Tanjung Uban, Kanwil Direktorat Jenderal Imigrasi Kepulauan Riau, " + tanggalFormal + " :\n\n";

    kegiatanHariIni.forEach((row, index) => {
      const namaKegiatan = row[headers.indexOf("NAMA_KEGIATAN")];
      textReportBackup += "*" + (index + 1) + ". " + namaKegiatan + "*\n";

      const platforms = [
        { nama: "Instagram", kolom: "LINK_INSTAGRAM" },
        { nama: "Facebook", kolom: "LINK_FACEBOOK" },
        { nama: "X (Twitter)", kolom: "LINK_X" },
        { nama: "Tiktok", kolom: "LINK_TIKTOK" },
        { nama: "Youtube Shorts", kolom: "LINK_YT_SHORTS" },
        { nama: "Youtube", kolom: "LINK_YOUTUBE" },
        { nama: "Laman Instansi", kolom: "LINK_INSTANSI" },
        { nama: "Pemberitaan Media", kolom: "LINK_MEDIA" }
      ];

      platforms.forEach(p => {
        const url = row[headers.indexOf(p.kolom)];
        if (url && url.toString().trim() !== "") {
          textReportBackup += p.nama + ":\n" + url + "\n";
        }
      });
      textReportBackup += "\n";
    });

    textReportBackup += "Demikian laporan ini disampaikan, selanjutnya mohon petunjuk dan arahan pimpinan.\n\n" +
                        "Tanjung Uban, " + new Date().toLocaleDateString('id-ID', { year: 'numeric', month: 'long', day: 'numeric' }) + "\n" +
                        "Kepala Kantor\n\n" +
                        "Adi Hari Pianto";

    // Panggil Gemini API
    const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey) return textReportBackup;

    const paketDataMentah = kegiatanHariIni.map((row, index) => {
      return {
        nomor_urut: index + 1,
        nama_kegiatan: row[headers.indexOf("NAMA_KEGIATAN")],
        links: {
          "Instagram": row[headers.indexOf("LINK_INSTAGRAM")],
          "Facebook": row[headers.indexOf("LINK_FACEBOOK")],
          "X (Twitter)": row[headers.indexOf("LINK_X")],
          "Tiktok": row[headers.indexOf("LINK_TIKTOK")],
          "Youtube Shorts": row[headers.indexOf("LINK_YT_SHORTS")],
          "Youtube": row[headers.indexOf("LINK_YOUTUBE")],
          "Laman Instansi": row[headers.indexOf("LINK_INSTANSI")],
          "Pemberitaan Media": row[headers.indexOf("LINK_MEDIA")]
        }
      };
    });

    const urlAPI = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;
    const promptSistem = "Kamu adalah sistem otomatisasi pelaporan naskah dinas formal terstruktur untuk pimpinan tinggi. Tugasmu menyusun draf laporan Atensi Pimpinan berdasarkan data mentah JSON.\n" +
                         "STRUKTUR WAJIB:\n1. Mulai dengan 'Assalamu’alaikum Wr. Wb.'\n2. Cetak daftar Yth pimpinan (1. Direktur Jenderal Imigrasi;, 2. Sekretaris Direktorat Jenderal Imigrasi;, 3. Para Direktur di lingkungan Direktorat Jenderal Imigrasi;, 4. Kepala Kantor Wilayah Direktorat Jenderal Imigrasi Kepulauan Riau;).\n3. Dari: 'Kepala Kantor Imigrasi Kelas II TPI Tanjung Uban'.\n4. Pengantar: 'Bersama ini dengan hormat melaporkan terkait Publikasi Kantor Imigrasi Kelas II TPI Tanjung Uban, Kanwil Direktorat Jenderal Imigrasi Kepulauan Riau...'\n5. Setiap kegiatan dicetak format Bold bintang tunggal WhatsApp: *1. Nama Kegiatan*.\n6. DILARANG KERAS memakai double/triple bintang (*** / **).\n7. Tampilkan nama platform dan link tepat di bawahnya jika link tersedia. Lewati jika kosong.\n8. JANGAN tampilkan statistik atau ringkasan apapun di akhir naskah.\n9. Tutup dengan 'Demikian laporan ini disampaikan, selanjutnya mohon petunjuk dan arahan pimpinan.', nama kota, jabatan 'Kepala Kantor', dan nama 'Adi Hari Pianto'.";

    const payload = {
      contents: [{ parts: [{ text: promptSistem + "\n\nData JSON:\n" + JSON.stringify(paketDataMentah) }] }]
    };

    const opsiRequest = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const respon = UrlFetchApp.fetch(urlAPI, opsiRequest);
    const hasilJson = JSON.parse(respon.getContentText());
    
    if (hasilJson.candidates && hasilJson.candidates[0].content.parts[0].text) {
      return hasilJson.candidates[0].content.parts[0].text;
    } else {
      return textReportBackup;
    }
  } catch (error) {
    return textReportBackup;
  }
}