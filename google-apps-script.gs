const SHEET_NAME = 'Sparepart';
const HEADERS = ['id', 'nama', 'stok', 'satuan', 'harga', 'updatedAt'];

function comparableValue(value, columnIndex) {
  if (columnIndex === 5) {
    if (value instanceof Date) return value.getTime();
    const match = String(value || '').match(/^(\d{2})\/(\d{2})\/(\d{2,4}) (\d{2}):(\d{2})/);
    if (match) {
      const year = match[3].length === 2 ? 2000 + Number(match[3]) : Number(match[3]);
      return Date.UTC(year, Number(match[2]) - 1, Number(match[1]), Number(match[4]) - 7, Number(match[5]));
    }
  }
  return String(value ?? '').trim();
}

function rowsAreEqual(currentRow, incomingRow) {
  return HEADERS.every((header, index) => comparableValue(currentRow[index], index) === comparableValue(incomingRow[index], index));
}

function doPost(request) {
  const payload = JSON.parse(request.postData.contents || '{}');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || payload.action !== 'upsert') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: 'Konfigurasi tidak valid.' })).setMimeType(ContentService.MimeType.JSON);
  }

  const rows = payload.data || [];
  const lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    if (HEADERS.some((header, index) => currentHeaders[index] !== header)) {
      const currentRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues() : [];
      const headerIndexes = new Map(currentHeaders.map((header, index) => [header, index]));
      const migratedRows = currentRows.map(row => HEADERS.map(header => row[headerIndexes.get(header)] ?? ''));
      sheet.getRange(1, 1, Math.max(1, migratedRows.length + 1), HEADERS.length).setValues([HEADERS, ...migratedRows]);
    }
  }

  const existingRows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues()
    : [];
  const rowById = new Map();
  existingRows.forEach((row, index) => {
    if (row[0]) rowById.set(String(row[0]), index + 2);
  });

  const incomingIds = new Set();
  const rowsToDelete = [];
  existingRows.forEach((row, index) => {
    if (!rows.some(item => String(item.id) === String(row[0]))) rowsToDelete.push(index + 2);
  });

  rows.forEach(item => {
    const values = [item.id, item.nama, item.stok, item.satuan, item.harga, item.updatedAt];
    const rowNumber = rowById.get(String(item.id));
    incomingIds.add(String(item.id));

    if (rowNumber) {
      const currentValues = sheet.getRange(rowNumber, 1, 1, HEADERS.length).getValues()[0];
      if (!rowsAreEqual(currentValues, values)) {
        sheet.getRange(rowNumber, 1, 1, HEADERS.length).setValues([values]);
      }
    } else {
      const newRow = sheet.getLastRow() + 1;
      sheet.getRange(newRow, 1, 1, HEADERS.length).setValues([values]);
      rowById.set(String(item.id), newRow);
    }
  });

  rowsToDelete.reverse().forEach(rowNumber => sheet.deleteRow(rowNumber));
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
