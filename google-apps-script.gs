const SHEET_NAME = 'Sparepart';

function doPost(request) {
  const payload = JSON.parse(request.postData.contents);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || payload.action !== 'upsert') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: 'Konfigurasi tidak valid.' })).setMimeType(ContentService.MimeType.JSON);
  }

  const rows = payload.data || [];
  sheet.clearContents();
  sheet.appendRow(['id', 'nama', 'jenis', 'merk', 'harga', 'stok', 'updatedAt']);
  rows.forEach(item => sheet.appendRow([item.id, item.nama, item.jenis, item.merk, item.harga, item.stok, item.updatedAt]));
  return ContentService.createTextOutput(JSON.stringify({ ok: true, count: rows.length })).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  const data = values.map(row => headers.reduce((record, header, index) => ({ ...record, [header]: row[index] }), {}));
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
