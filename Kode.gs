// ID Spreadsheet Anda (diambil dari URL yang Anda berikan)
const SPREADSHEET_ID = '14ZHt1cy02hV4sj6Sjuxq-xvMgdIqWPiB';

// Fungsi untuk menampilkan halaman HTML
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Kuesioner CSI IT - PT. Havindo Pakan Optima')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Fungsi untuk mengecek apakah divisi sudah mengisi pada quarter tersebut
function checkExistingData(divisi, quarter) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = "Quarter " + quarter; 
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return { exists: false };
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][1] === divisi) {
      return { 
        exists: true, 
        row: i + 1, 
        message: "Divisi " + divisi + " sudah mengisi kuesioner untuk Quarter " + quarter + "." 
      };
    }
  }
  return { exists: false };
}

// Fungsi untuk menyimpan atau mengupdate data kuesioner dan saran
function submitKuesioner(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = "Quarter " + formData.quarter;
    let sheet = ss.getSheetByName(sheetName);
    
    // Jika sheet belum ada, buatkan otomatis
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow([
        "Timestamp", "Divisi", 
        "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9", "P10", "P11", 
        "Total Skor", "Skor CSI %"
      ]);
    }
    
    const timestamp = new Date();
    const totalScore = parseInt(formData.q1) + parseInt(formData.q2) + parseInt(formData.q3) + 
                       parseInt(formData.q4) + parseInt(formData.q5) + parseInt(formData.q6) + 
                       parseInt(formData.q7) + parseInt(formData.q8) + parseInt(formData.q9) + 
                       parseInt(formData.q10) + parseInt(formData.q11);
    
    const maxScore = 55; // 11 pertanyaan x 5 poin maksimal
    const csiPercentage = (totalScore / maxScore);
    
    const rowData = [
      timestamp,
      formData.divisi,
      formData.q1,
      formData.q2,
      formData.q3,
      formData.q4,
      formData.q5,
      formData.q6,
      formData.q7,
      formData.q8,
      formData.q9,
      formData.q10,
      formData.q11,
      totalScore,
      csiPercentage
    ];

    let messageOutput = "";

    if (formData.action === 'update' && formData.updateRow) {
      const range = sheet.getRange(formData.updateRow, 1, 1, rowData.length);
      range.setValues([rowData]);
      messageOutput = "Data berhasil diperbarui!";
    } else {
      sheet.appendRow(rowData);
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 15).setNumberFormat("0.00%");
      messageOutput = "Terima kasih! Kuesioner berhasil dikirim.";
    }
    
    // ====================================================
    // HANDLE TAB "Saran dan kritik"
    // ====================================================
    const saranSheetName = "Saran dan kritik";
    let saranSheet = ss.getSheetByName(saranSheetName);
    
    if (!saranSheet) {
      saranSheet = ss.insertSheet(saranSheetName);
      saranSheet.appendRow(["Timestamp", "Divisi", "Quarter", "Saran & Kritik", "Keterangan Terakhir"]);
      saranSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#e0f2fe");
    }
    
    const strTime = Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss");
    const isiSaran = (formData.saran && formData.saran.trim() !== "") ? formData.saran : "-";
    
    if (formData.action === 'update' && formData.updateRow) {
      const saranData = saranSheet.getDataRange().getValues();
      let foundSaranRow = -1;
      
      for (let i = saranData.length - 1; i > 0; i--) {
        if (saranData[i][1] === formData.divisi && String(saranData[i][2]) === String(formData.quarter)) {
          foundSaranRow = i + 1;
          break;
        }
      }
      
      if (foundSaranRow !== -1) {
        saranSheet.getRange(foundSaranRow, 1).setValue(timestamp);
        saranSheet.getRange(foundSaranRow, 4).setValue(isiSaran);
        saranSheet.getRange(foundSaranRow, 5).setValue("Diedit pada: " + strTime);
      } else {
        saranSheet.appendRow([timestamp, formData.divisi, formData.quarter, isiSaran, "Ditambahkan saat edit: " + strTime]);
      }
    } else {
      saranSheet.appendRow([timestamp, formData.divisi, formData.quarter, isiSaran, "Input Baru"]);
    }

    return { success: true, message: messageOutput };
    
  } catch (error) {
    return { success: false, message: "Terjadi kesalahan: " + error.toString() };
  }
}
