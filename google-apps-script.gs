const SHEET_NAME = 'Sparepart';

function doPost(request) {
  const payload = JSON.parse(request.postData.contents || '{}');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || payload.action !== 'upsert') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: 'Konfigurasi tidak valid.' })).setMimeType(ContentService.MimeType.JSON);
  }

  const rows = payload.data || [];
  sheet.clearContents();
  sheet.appendRow(['id', 'nama', 'stok', 'satuan', 'harga', 'updatedAt']);
  rows.forEach(item => sheet.appendRow([item.id, item.nama, item.stok, item.satuan, item.harga, item.updatedAt]));
  return ContentService.createTextOutput(JSON.stringify({ ok: true, count: rows.length })).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: 'Tab Sparepart tidak ditemukan.' })).setMimeType(ContentService.MimeType.JSON);
  }

  const values = sheet.getDataRange().getValues();
  if (values.length === 0) {
    return ContentService.createTextOutput(JSON.stringify({ ok: true, data: [], revision: '' })).setMimeType(ContentService.MimeType.JSON);
  }

  const headers = values.shift();
  const data = values.map(row => headers.reduce((record, header, index) => ({ ...record, [header]: row[index] }), {}));
  const revision = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(data)));
  return ContentService.createTextOutput(JSON.stringify({ ok: true, data: data, revision: revision })).setMimeType(ContentService.MimeType.JSON);
}
