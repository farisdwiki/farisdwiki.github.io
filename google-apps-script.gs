const SHEET_NAME = 'Sparepart';
const HEADERS = ['id', 'nama', 'stok', 'satuan', 'harga', 'updatedAt'];

function normalizeHeader(value) {
  return String(value || '').trim().toLowerCase();
}

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

function ensureHeaders(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(2, 6, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('dd/MM/yyyy HH:mm');
    return;
  }

  const columnCount = Math.max(HEADERS.length, sheet.getLastColumn());
  const currentHeaders = sheet.getRange(1, 1, 1, columnCount).getValues()[0];
  const normalizedHeaders = currentHeaders.map(normalizeHeader);
  const headersMatch = HEADERS.length === currentHeaders.length && HEADERS.every((header, index) => normalizedHeaders[index] === header);
  if (!headersMatch) {
    const currentRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, columnCount).getValues() : [];
    const headerIndexes = new Map(normalizedHeaders.map((header, index) => [header, index]));
    const migratedRows = currentRows.map(row => HEADERS.map(header => row[headerIndexes.get(header)] ?? ''));
    sheet.getRange(1, 1, Math.max(1, migratedRows.length + 1), HEADERS.length).setValues([HEADERS, ...migratedRows]);
    if (sheet.getLastColumn() > HEADERS.length) sheet.deleteColumns(HEADERS.length + 1, sheet.getLastColumn() - HEADERS.length);
  }
  if (sheet.getLastRow() > 1) sheet.getRange(2, 6, sheet.getLastRow() - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
}

function findRowById(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const index = ids.findIndex(row => String(row[0]) === String(id));
  return index < 0 ? 0 : index + 2;
}

function normalizeUpdatedAt(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return formatUpdatedAt(value);
  const text = String(value || '').trim();
  const match = text.match(/^(\d{2})\/(\d{2})\/(\d{2,4}) (\d{2}):(\d{2})(?:\s*WIB)?$/i);
  if (match) {
    const year = match[3].length === 2 ? 2000 + Number(match[3]) : Number(match[3]);
    return formatUpdatedAt(new Date(Date.UTC(year, Number(match[2]) - 1, Number(match[1]), Number(match[4]) - 7, Number(match[5]))));
  }
  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? formatUpdatedAt(new Date()) : formatUpdatedAt(parsed);
}

function formatUpdatedAt(value) {
  const date = value instanceof Date ? value : new Date(value);
  return Utilities.formatDate(date, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm') + ' WIB';
}

function itemValues(item, updateTimestamp = false) {
  const updatedAt = updateTimestamp ? formatUpdatedAt(new Date()) : normalizeUpdatedAt(item.updatedAt);
  return [item.id, item.nama, Number(item.stok || 0), item.satuan || 'pcs', Number(item.harga || 0), updatedAt];
}

function writeRow(sheet, rowNumber, values) {
  sheet.getRange(rowNumber, 1, 1, HEADERS.length).setValues([values]);
  sheet.getRange(rowNumber, 6).setNumberFormat('@');
  sheet.getRange(rowNumber, 6).setValue(String(values[5] || formatUpdatedAt(new Date())));
  SpreadsheetApp.flush();
}

function handlePost(request) {
  const payload = JSON.parse(request.postData.contents || '{}');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || payload.action !== 'upsert') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: 'Konfigurasi tidak valid.' })).setMimeType(ContentService.MimeType.JSON);
  }

  ensureHeaders(sheet);

  if (payload.mode === 'single') {
    const item = payload.item || {};
    if (!String(item.id || '').trim()) return ContentService.createTextOutput(JSON.stringify({ ok: false, message: 'ID data wajib diisi.' })).setMimeType(ContentService.MimeType.JSON);
    const rowNumber = findRowById(sheet, item.id);

    if (payload.operation === 'delete') {
      if (rowNumber) sheet.deleteRow(rowNumber);
      return ContentService.createTextOutput(JSON.stringify({ ok: true, operation: 'delete' })).setMimeType(ContentService.MimeType.JSON);
    }

    const values = itemValues(item, true);
    if (rowNumber) {
      const currentValues = sheet.getRange(rowNumber, 1, 1, HEADERS.length).getValues()[0];
      if (!rowsAreEqual(currentValues, values)) writeRow(sheet, rowNumber, values);
    } else {
      writeRow(sheet, sheet.getLastRow() + 1, values);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true, operation: payload.operation || 'upsert', updatedAt: values[5] })).setMimeType(ContentService.MimeType.JSON);
  }

  const rows = payload.data || [];
  const dataLastRow = sheet.getLastRow();
  const existingRows = dataLastRow > 1
    ? sheet.getRange(2, 1, dataLastRow - 1, HEADERS.length).getValues()
    : [];
  const rowById = new Map();
  existingRows.forEach((row, index) => {
    if (row[0]) rowById.set(String(row[0]), index + 2);
  });

  const incomingIds = new Set(rows.map(item => String(item.id)));
  const rowsToDelete = [];
  existingRows.forEach((row, index) => {
    if (!incomingIds.has(String(row[0]))) rowsToDelete.push(index + 2);
  });

  const newRows = [];
  rows.forEach(item => {
    const values = itemValues(item, true);
    const rowNumber = rowById.get(String(item.id));

    if (rowNumber) {
      const currentValues = existingRows[rowNumber - 2];
      if (!rowsAreEqual(currentValues, values)) {
        writeRow(sheet, rowNumber, values);
      }
    } else {
      newRows.push(values);
    }
  });

  if (newRows.length) {
    const firstNewRow = sheet.getLastRow() + 1;
    sheet.getRange(firstNewRow, 1, newRows.length, HEADERS.length).setValues(newRows);
    for (let index = 0; index < newRows.length; index++) {
      sheet.getRange(firstNewRow + index, 6).setNumberFormat('@');
      sheet.getRange(firstNewRow + index, 6).setValue(String(newRows[index][5] || formatUpdatedAt(new Date())));
    }
    SpreadsheetApp.flush();
  }

  rowsToDelete.reverse().forEach(rowNumber => sheet.deleteRow(rowNumber));
  return ContentService.createTextOutput(JSON.stringify({ ok: true, count: rows.length })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(request) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return handlePost(request);
  } finally {
    lock.releaseLock();
  }
}

function handleGet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: 'Tab Sparepart tidak ditemukan.' })).setMimeType(ContentService.MimeType.JSON);
  }

  ensureHeaders(sheet);
  const values = sheet.getDataRange().getValues();
  if (values.length === 0) {
    return ContentService.createTextOutput(JSON.stringify({ ok: true, data: [], revision: '' })).setMimeType(ContentService.MimeType.JSON);
  }

  const headers = values.shift();
  const data = values.map(row => headers.reduce((record, header, index) => ({ ...record, [header]: row[index] }), {}));
  const revision = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(data)));
  return ContentService.createTextOutput(JSON.stringify({ ok: true, data: data, revision: revision })).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return handleGet();
  } finally {
    lock.releaseLock();
  }
}
