// Google Apps Script for Irma & Rio digital invitation
// Spreadsheet ID: 1c3h26lDVyFdRC1wnvnTUaETTj5-FfXg6iYV2FnlH5XI
// Sheet Name: Sheet1

const SPREADSHEET_ID = "1c3h26lDVyFdRC1wnvnTUaETTj5-FfXg6iYV2FnlH5XI";
const SHEET_NAME = "Sheet1";

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const rows = sheet.getDataRange().getValues();
    const data = [];
    
    // Check if sheet is empty or only has headers
    if (rows.length > 1) {
      const headers = rows[0];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        
        // Format timestamp beautifully
        let formattedDate = "";
        if (row[0] instanceof Date) {
          formattedDate = row[0].toLocaleDateString("id-ID", {
            day: "numeric",
            month: "short",
            year: "numeric",
            hour: "2-digit",
            minute: "2-digit"
          });
        } else {
          formattedDate = row[0] ? row[0].toString() : "";
        }

        data.push({
          timestamp: formattedDate,
          nama: row[1] ? row[1].toString() : "",
          ucapan: row[2] ? row[2].toString() : "",
          kehadiran: row[3] ? row[3].toString() : "Hadir",
          jumlah: row[4] ? row[4].toString() : "1"
        });
      }
    }
    
    // Return JSON output with CORS support
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    let params;
    if (e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else {
      params = e.parameter;
    }
    
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    
    // Create headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "Nama Tamu", "Ucapan", "Konfirmasi Kehadiran", "Jumlah Tamu"]);
    }
    
    const timestamp = new Date();
    const nama = params.nama || "";
    const ucapan = params.ucapan || "";
    const kehadiran = params.kehadiran || "";
    const jumlah = params.jumlah || "";
    
    sheet.appendRow([timestamp, nama, ucapan, kehadiran, jumlah]);
    
    return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
