/**
 * Logistics Suite — Fuel Tracker Backend (v3)
 * Per-Vehicle Efficiency Analyzer + Token-based Admin Auth
 *
 * Spreadsheet ID: 1GJLqBSsf1r963zX6t-EjlfH7l0xEXaI0R_qKZTj9lfA
 * Sheet Name:     Trips
 *
 * Auth flow:
 *   POST { action:'login', user, pass } → { status:'ok', token, ttl }
 *   GET  ?action=read&token=...        → array (token required)
 *   GET  ?action=logout&token=...      → invalidate token
 *
 * Public endpoints:
 *   GET  ?action=ping                  → health check
 *   POST { type, plate, ... }          → append trip (open to drivers)
 */

const SHEET_ID   = '1GJLqBSsf1r963zX6t-EjlfH7l0xEXaI0R_qKZTj9lfA';
const SHEET_NAME = 'Trips';
const REQUIRED_COLS = ['Timestamp','Type','Plate','Owner','Driver','JobOrder','Origin','Destination','StartKm','EndKm','Distance','FuelLiters','FuelCost'];

// ===== Admin credentials (server-side only — never sent to frontend) =====
const ADMIN_USER = 'admin';
const ADMIN_PASS = '007';
const TOKEN_TTL_SEC = 8 * 60 * 60; // 8 hours
const FAILED_LOGIN_DELAY_MS = 800;

/* ============== Utilities ============== */
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function normKey(s) {
  return String(s || '').toLowerCase().replace(/\s+/g, '');
}

/* ============== Auth helpers ============== */
function makeToken_() {
  return Utilities.getUuid().replace(/-/g, '');
}
function saveToken_(token) {
  CacheService.getScriptCache().put('tok_' + token, '1', TOKEN_TTL_SEC);
}
function isTokenValid_(token) {
  if (!token || typeof token !== 'string' || token.length < 8) return false;
  return !!CacheService.getScriptCache().get('tok_' + token);
}
function invalidateToken_(token) {
  if (token) CacheService.getScriptCache().remove('tok_' + token);
}

/* ============== Sheet helpers ============== */
function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, REQUIRED_COLS.length).setValues([REQUIRED_COLS]).setFontWeight('bold');
    return sh;
  }
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const existing = headerRow.map(normKey);
  let nextCol = lastCol + 1;
  REQUIRED_COLS.forEach(col => {
    if (existing.indexOf(normKey(col)) === -1) {
      sh.getRange(1, nextCol).setValue(col).setFontWeight('bold');
      nextCol++;
    }
  });
  return sh;
}

function getColIndex_(sh) {
  const lastCol = sh.getLastColumn();
  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = {};
  headerRow.forEach((h, i) => { idx[normKey(h)] = i + 1; });
  return idx;
}

/* ============== doGet ============== */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'read';
    const token = e && e.parameter && e.parameter.token;

    if (action === 'ping') {
      return jsonOut({ status: 'ok', message: 'pong' });
    }

    if (action === 'logout') {
      invalidateToken_(token);
      return jsonOut({ status: 'ok', message: 'logged out' });
    }

    if (action === 'verify') {
      return jsonOut({ status: isTokenValid_(token) ? 'ok' : 'unauthorized' });
    }

    if (action === 'read') {
      if (!isTokenValid_(token)) {
        return jsonOut({ status: 'unauthorized', message: 'login required' });
      }
      const sh = getSheet_();
      const lastRow = sh.getLastRow();
      const lastCol = sh.getLastColumn();
      if (lastRow < 2) return jsonOut([]);
      const data = sh.getRange(1, 1, lastRow, lastCol).getValues();
      const headers = data[0];
      const rows = data.slice(1)
        .filter(r => r.some(c => c !== '' && c !== null))
        .map(r => {
          const o = {};
          headers.forEach((h, i) => { o[normKey(h)] = r[i]; });
          if (o.timestamp instanceof Date) o.timestamp = o.timestamp.toISOString();
          return o;
        });
      return jsonOut(rows);
    }

    return jsonOut({ status: 'error', message: 'unknown action' });

  } catch (err) {
    return jsonOut({ status: 'error', message: String(err) });
  }
}

/* ============== doPost ============== */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut({ status: 'error', message: 'No postData' });
    }
    const body = JSON.parse(e.postData.contents);

    // -------- Login --------
    if (body.action === 'login') {
      const u = String(body.user || '').trim();
      const p = String(body.pass || '');
      if (u === ADMIN_USER && p === ADMIN_PASS) {
        const token = makeToken_();
        saveToken_(token);
        return jsonOut({ status: 'ok', token: token, ttl: TOKEN_TTL_SEC });
      }
      Utilities.sleep(FAILED_LOGIN_DELAY_MS);
      return jsonOut({ status: 'unauthorized', message: 'invalid credentials' });
    }

    // -------- Save trip (public — drivers can submit without login) --------
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);

      const type        = String(body.type || '').trim();
      const plate       = String(body.plate || '').trim();
      const owner       = String(body.owner || 'บริษัท').trim();
      const driver      = String(body.driver || '').trim();
      const jobOrder    = String(body.jobOrder || '').trim();
      const origin      = String(body.origin || '').trim();
      const destination = String(body.destination || '').trim();
      const startKm     = Number(body.startKm) || 0;
      const endKm       = Number(body.endKm) || 0;
      const distance    = Number(body.distance) || Math.max(0, endKm - startKm);
      const fuelLiters  = Number(body.fuelLiters) || 0;
      const fuelCost    = Number(body.fuelCost) || 0;

      if (!type || !plate || !driver || !jobOrder || !origin || !destination) {
        return jsonOut({ status: 'error', message: 'Missing required fields (type/plate/driver/jobOrder/origin/destination)' });
      }

      const sh = getSheet_();
      const idx = getColIndex_(sh);
      const lastCol = sh.getLastColumn();
      const row = new Array(lastCol).fill('');

      const fieldMap = {
        timestamp:   new Date(),
        type:        type,
        plate:       plate,
        owner:       owner,
        driver:      driver,
        joborder:    jobOrder,
        origin:      origin,
        destination: destination,
        startkm:     startKm,
        endkm:       endKm,
        distance:    distance,
        fuelliters:  fuelLiters,
        fuelcost:    fuelCost
      };

      Object.keys(fieldMap).forEach(key => {
        const col = idx[key];
        if (col) row[col - 1] = fieldMap[key];
      });

      sh.appendRow(row);
      return jsonOut({ status: 'ok', message: 'saved' });

    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }

  } catch (err) {
    return jsonOut({ status: 'error', message: String(err) });
  }
}

/* ============== Manual tests ============== */
function test_login() {
  const fake = { postData: { contents: JSON.stringify({ action:'login', user:'admin', pass:'007' }) } };
  Logger.log(doPost(fake).getContent());
}
function test_loginBad() {
  const fake = { postData: { contents: JSON.stringify({ action:'login', user:'admin', pass:'wrong' }) } };
  Logger.log(doPost(fake).getContent());
}
function test_readNoToken() {
  Logger.log(doGet({ parameter: { action:'read' } }).getContent());
}
