const SHEET_NAME = 'Sparepart';
const HEADERS = ['id', 'nama', 'stok', 'satuan', 'harga', 'updatedAt'];
const DATE_FORMAT = 'dd/MM/yyyy HH:mm';

function jsonResponse(body) {
  return ContentService.createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeHeader(value) {
  return String(value || '').trim().toLowerCase();
}

function getSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Tab '${SHEET_NAME}' tidak ditemukan.`);
  prepareSheet(sheet);
  return sheet;
}

function prepareSheet(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  } else {
    const width = Math.max(sheet.getLastColumn(), HEADERS.length);
    const currentHeaders = sheet.getRange(1, 1, 1, width).getValues()[0].map(normalizeHeader);
    const hasExpectedHeaders = HEADERS.every((header, index) => currentHeaders[index] === header);

    if (!hasExpectedHeaders || width !== HEADERS.length) {
      const oldRows = sheet.getLastRow() > 1
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues()
        : [];
      const indexes = new Map(currentHeaders.map((header, index) => [header, index]));
      const migratedRows = oldRows.map(row => HEADERS.map(header => {
        const index = indexes.get(header);
        return index === undefined ? '' : row[index];
      }));

      sheet.getRange(1, 1, migratedRows.length + 1, HEADERS.length)
        .setValues([HEADERS, ...migratedRows]);
      if (sheet.getMaxColumns() > HEADERS.length) {
        sheet.deleteColumns(HEADERS.length + 1, sheet.getMaxColumns() - HEADERS.length);
      }
    }
  }

  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 6, sheet.getMaxRows() - 1, 1).setNumberFormat(DATE_FORMAT);
  }
}

function parseDate(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const text = String(value || '').trim();
  const match = text.match(/^(\d{2})\/(\d{2})\/(\d{2,4})[ T](\d{2}):(\d{2})(?::(\d{2}))?(?:\s*WIB)?$/i);
  if (match) {
    const year = match[3].length === 2 ? 2000 + Number(match[3]) : Number(match[3]);
    return new Date(Date.UTC(year, Number(match[2]) - 1, Number(match[1]), Number(match[4]) - 7, Number(match[5]), Number(match[6] || 0)));
  }
  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function dateForItem(value) {
  return parseDate(value) || new Date();
}

function normalizeItem(item, fallbackDate) {
  const source = item || {};
  return [
    String(source.id || '').trim(),
    String(source.nama || '').trim(),
    Number(source.stok || 0),
    String(source.satuan || 'pcs').trim() || 'pcs',
    Number(source.harga || 0),
    dateForItem(source.updatedAt || fallbackDate)
  ];
}

function findRow(sheet, id) {
  if (!id || sheet.getLastRow() < 2) return 0;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  const index = ids.findIndex(row => String(row[0]).trim() === String(id).trim());
  return index < 0 ? 0 : index + 2;
}

function writeRow(sheet, rowNumber, values) {
  sheet.getRange(rowNumber, 1, 1, HEADERS.length).setValues([values]);
  sheet.getRange(rowNumber, 6).setNumberFormat(DATE_FORMAT);
}

function handleSingle(sheet, payload) {
  const item = payload.item || {};
  const id = String(item.id || '').trim();
  if (!id) return jsonResponse({ ok: false, message: 'ID data wajib diisi.' });

  const rowNumber = findRow(sheet, id);
  if (payload.operation === 'delete') {
    if (rowNumber) sheet.deleteRow(rowNumber);
    return jsonResponse({ ok: true, operation: 'delete' });
  }

  const values = normalizeItem(item, new Date());
  writeRow(sheet, rowNumber || sheet.getLastRow() + 1, values);
  SpreadsheetApp.flush();
  return jsonResponse({ ok: true, operation: payload.operation || 'upsert', updatedAt: values[5].toISOString() });
}

function handleSync(sheet, payload) {
  const incoming = Array.isArray(payload.data) ? payload.data : [];
  const lastRow = sheet.getLastRow();
  const existing = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues() : [];
  const rowsById = new Map();
  existing.forEach((row, index) => {
    if (row[0]) rowsById.set(String(row[0]).trim(), index + 2);
  });

  const incomingIds = new Set(incoming.map(item => String(item && item.id || '').trim()).filter(Boolean));
  const rowsToDelete = existing
    .map((row, index) => ({ id: String(row[0] || '').trim(), rowNumber: index + 2 }))
    .filter(entry => entry.id && !incomingIds.has(entry.id))
    .map(entry => entry.rowNumber)
    .reverse();

  incoming.forEach(item => {
    const id = String(item && item.id || '').trim();
    if (!id) return;
    const values = normalizeItem(item, new Date());
    writeRow(sheet, rowsById.get(id) || sheet.getLastRow() + 1, values);
  });

  rowsToDelete.forEach(rowNumber => sheet.deleteRow(rowNumber));
  SpreadsheetApp.flush();
  return jsonResponse({ ok: true, count: incoming.length });
}

function handlePost(request) {
  const payload = JSON.parse(request.postData && request.postData.contents || '{}');
  if (payload.action !== 'upsert') return jsonResponse({ ok: false, message: 'Action tidak valid.' });
  const sheet = getSheet();
  return payload.mode === 'single' ? handleSingle(sheet, payload) : handleSync(sheet, payload);
}

function handleGet() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  const data = values.slice(1).filter(row => row[0]).map(row => ({
    id: row[0],
    nama: row[1],
    stok: row[2],
    satuan: row[3],
    harga: row[4],
    updatedAt: row[5] instanceof Date ? row[5].toISOString() : row[5]
  }));
  const revision = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(data))
  );
  return jsonResponse({ ok: true, data: data, revision: revision });
}

function doPost(request) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return handlePost(request);
  } catch (error) {
    return jsonResponse({ ok: false, message: error.message });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return handleGet();
  } catch (error) {
    return jsonResponse({ ok: false, message: error.message });
  } finally {
    lock.releaseLock();
  }
}
