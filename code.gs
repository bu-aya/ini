/**
 * SISTEM ADMINISTRASI SDIT ANNIHAYAH - BACKEND ENGINE
 * Fitur: Auto-Setup Database & CRUD API
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();

// Fungsi Setup Database Otomatis
function setupDatabase() {
  const sheets = {
    "Siswa": ["id", "nama", "nis", "nisn", "gender", "tempatLahir", "tanggalLahir", "namaAyah", "namaIbu", "kelas", "status", "tahunAjaran", "noIjazah", "tanggalLulus", "nilaiRataIjazah"],
    "Nilai": ["id", "siswaId", "kelas", "semester", "tahunAjaran", "PAI", "PPKN", "BIndo", "MTK", "IPA", "IPS", "IPAS", "SBDP", "PJOK", "BSunda", "BInggris", "sakit", "izin", "alfa", "catatan"],
    "SuratKeluar": ["id", "tipe", "noSurat", "tanggalKirim", "penerima", "keterangan", "siswaId"],
    "CPTP": ["kelas", "mapel", "cp", "tp"]
  };

  for (let name in sheets) {
    let sheet = SS.getSheetByName(name);
    if (!sheet) {
      sheet = SS.insertSheet(name);
      sheet.getRange(1, 1, 1, sheets[name].length).setValues([sheets[name]]).setFontWeight("bold").setBackground("#f3f3f3");
    }
  }
  return "Database Berhasil Disiapkan!";
}

// Handler GET: Mengambil seluruh data untuk UI
function doGet() {
  const data = {};
  const sheetNames = ["Siswa", "Nilai", "SuratKeluar", "CPTP"];
  
  sheetNames.forEach(name => {
    const sheet = SS.getSheetByName(name);
    if (sheet) {
      const vals = sheet.getDataRange().getValues();
      const headers = vals.shift();
      data[name] = vals.map(row => {
        let obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      });
    } else {
      data[name] = [];
    }
  });

  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// Handler POST: Menyimpan/Update Data
function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    const sheet = SS.getSheetByName(request.sheet);
    const payload = request.data;

    if (!sheet) throw new Error("Sheet tidak ditemukan");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    if (action === "save") {
      // Cari jika ID sudah ada (Update) atau baru (Append)
      const dataRange = sheet.getDataRange().getValues();
      let rowIndex = -1;
      
      if (payload.id) {
        for (let i = 1; i < dataRange.length; i++) {
          if (dataRange[i][0].toString() === payload.id.toString()) {
            rowIndex = i + 1;
            break;
          }
        }
      }

      const rowValues = headers.map(h => payload[h] !== undefined ? payload[h] : "");
      
      if (rowIndex !== -1) {
        sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
      } else {
        sheet.appendRow(rowValues);
      }
    } 
    
    else if (action === "delete") {
      const dataRange = sheet.getDataRange().getValues();
      for (let i = 1; i < dataRange.length; i++) {
        if (dataRange[i][0].toString() === payload.id.toString()) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}
