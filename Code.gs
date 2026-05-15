/**
 * Logistics Suite — Fuel Tracker Backend (v2)
 * Per-Vehicle Efficiency Analyzer
 *
 * Spreadsheet ID: 1GJLqBSsf1r963zX6t-EjlfH7l0xEXaI0R_qKZTj9lfA
 * Sheet Name:     Trips
 *
 * Dynamic header mapping — handles ANY column ordering / casing
 * Auto-adds 'Owner' column if missing
 */

const SHEET_ID   = '1GJLqBSsf1r963zX6t-EjlfH7l0xEXaI0R_qKZTj9lfA';
const SHEET_NAME = 'Trips';
const REQUIRED_COLS = ['Timestamp','Type','Plate','Owner','StartKm','EndKm','Distance','FuelLiters','FuelCost'];

/** JSON response helper */
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Normalize header text for matching (lowercase, strip spaces) */
function normKey(s) {
  return String(s || '').toLowerCase().replace(/\s+/g, '');
}

/** Open Trips sheet — auto-create + auto-migrate missing columns */
function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, REQUIRED_COLS.length).setValues([REQUIRED_COLS]).setFontWeight('bold');
    return sh;
  }
  // Read current header row
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const existing = headerRow.map(normKey);

  // Add any missing required columns at the end
  let nextCol = lastCol + 1;
  REQUIRED_COLS.forEach(col => {
    if (existing.indexOf(normKey(col)) === -1) {
      // Avoid clashing with Thai "ค่าน้ำมัน (บาท)" → still adds new English col
      sh.getRange(1, nextCol).setValue(col).setFontWeight('bold');
      nextCol++;
    }
  });
  return sh;
}

/** Build a lookup: normKey(header) → column index (1-based) */
function getColIndex_(sh) {
  const lastCol = sh.getLastColumn();
  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = {};
  headerRow.forEach((h, i) => { idx[normKey(h)] = i + 1; });
  return idx;
}

/**
 * doGet — return all rows as JSON array
 * URL: ...?action=read  or  ?action=ping
 */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'read';
    if (action === 'ping') return jsonOut({ status: 'ok', message: 'pong' });

    const sh = getSheet_();
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2) return jsonOut([]);

    const data = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = data[0];
    const rows = data.slice(1)
      .filter(r => r.some(c => c !== '' && c !== null))   // skip blank rows
      .map(r => {
        const o = {};
        headers.forEach((h, i) => { o[normKey(h)] = r[i]; });
        if (o.timestamp instanceof Date) o.timestamp = o.timestamp.toISOString();
        return o;
      });
    return jsonOut(rows);

  } catch (err) {
    return jsonOut({ status: 'error', message: String(err) });
  }
}

/**
 * doPost — append a new trip row using dynamic header lookup
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut({ status: 'error', message: 'No postData' });
    }
    const body = JSON.parse(e.postData.contents);

    // Coerce inputs
    const type       = String(body.type || '').trim();
    const plate      = String(body.plate || '').trim();
    const owner      = String(body.owner || 'บริษัท').trim();
    const startKm    = Number(body.startKm) || 0;
    const endKm      = Number(body.endKm) || 0;
    const distance   = Number(body.distance) || Math.max(0, endKm - startKm);
    const fuelLiters = Number(body.fuelLiters) || 0;
    const fuelCost   = Number(body.fuelCost) || 0;

    if (!type || !plate) {
      return jsonOut({ status: 'error', message: 'Missing type or plate' });
    }

    const sh = getSheet_();
    const idx = getColIndex_(sh);
    const lastCol = sh.getLastColumn();
    const row = new Array(lastCol).fill('');

    // Map fields to their column position (case-insensitive)
    const fieldMap = {
      timestamp:  new Date(),
      type:       type,
      plate:      plate,
      owner:      owner,
      startkm:    startKm,
      endkm:      endKm,
      distance:   distance,
      fuelliters: fuelLiters,
      fuelcost:   fuelCost
    };

    Object.keys(fieldMap).forEach(key => {
      const col = idx[key];
      if (col) row[col - 1] = fieldMap[key];
    });

    sh.appendRow(row);
    return jsonOut({ status: 'ok', message: 'saved' });

  } catch (err) {
    return jsonOut({ status: 'error', message: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/** Manual test — run from Apps Script editor */
function test_doPost() {
  const fake = {
    postData: {
      contents: JSON.stringify({
        type: '4W', plate: 'TEST-001', owner: 'บริษัท',
        startKm: 100, endKm: 250, distance: 150,
        fuelLiters: 18, fuelCost: 650
      })
    }
  };
  Logger.log(doPost(fake).getContent());
}

function test_doGet() {
  Logger.log(doGet({ parameter: { action: 'read' } }).getContent().substring(0, 600));
}
