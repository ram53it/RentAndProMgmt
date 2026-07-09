/**
 * Jewel + Rent Manager — Google Sheets gateway (Apps Script Web App)
 * ------------------------------------------------------------------
 * This is the ONLY thing that touches your Google Sheets. Deploy it once,
 * keep your sheets private (un-shared), and the app talks to this instead of
 * the Google Sheets API. No billing account is required — Apps Script is free.
 *
 * SETUP
 *  1. Open your Google Sheet → Extensions → Apps Script.
 *  2. Delete any sample code, paste this whole file.
 *  3. Change SECRET below to a long random string of your own.
 *  4. Deploy → New deployment → type "Web app".
 *       Execute as: Me      Who has access: Anyone
 *     Authorize when asked (this is a normal consent, NOT billing).
 *  5. Copy the Web app URL that ends in /exec.
 *  6. In the app's Setup, paste: both Sheet IDs, the /exec URL, this SECRET,
 *     and your Google OAuth Client ID (for the sign-in gate).
 *  7. Un-share both sheets so only you (and this script) can touch them.
 *
 * ACCESS CONTROL (two factors)
 *  - Whitelisted Google account: only emails in WHITELIST below may sign in.
 *    The app sends a Google ID token; this script verifies it and issues a
 *    signed 7-day session. Data calls need that session — the shared secret
 *    alone is NOT enough.
 *  - PIN: decrypts the app config on the device (separate, client-side).
 *
 * DATA
 *  - Columns are matched by NAME, so the app never needs a column in a fixed
 *    position; ensureSchema adds any missing tab/column (at the end) for you.
 *  - Read / append / update / delete are all allowed (access is still limited
 *    to whitelisted Google accounts holding a valid session).
 */

var SECRET = 'CHANGE-ME-to-a-long-random-string';

// Your Google OAuth *Client ID* (Web application) — same value you put in the app's Setup.
var CLIENT_ID = 'YOUR_CLIENT_ID.apps.googleusercontent.com';

// ONLY these Google accounts may use the app. Add/remove emails here anytime.
var WHITELIST = [
  'you@gmail.com',
  // 'brother@gmail.com',
];

var SESSION_DAYS = 7;

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    if (req.secret !== SECRET) return _json({ ok: false, error: 'Bad secret' });

    // Step 1 — exchange a Google ID token for a signed session (whitelist check).
    if (req.action === 'login') return _json(_login(req.idToken));

    // Every other action requires a valid, unexpired, whitelisted session.
    var sess = _checkSession(req.session);
    if (!sess.ok) return _json({ ok: false, needLogin: true, error: 'Sign in with your Google account again.' });

    switch (req.action) {
      case 'read':         return _json({ ok: true, values: _read(req.spreadsheetId, req.sheet) });
      case 'readAll':      return _json({ ok: true, sheets: _readAll(req.spreadsheetId, req.sheets) });
      case 'append':       return _json({ ok: true, row: _append(req.spreadsheetId, req.sheet, req.values) });
      case 'update':       return _json({ ok: true, updated: _update(req.spreadsheetId, req.sheet, req.row, req.values) });
      case 'delete':       return _json({ ok: true, deleted: _delete(req.spreadsheetId, req.sheet, req.row, req.field, req.value) });
      case 'ensureSchema': return _json({ ok: true, changed: _ensureSchema(req.spreadsheetId, req.schema) });
      default:             return _json({ ok: false, error: 'Unknown action: ' + req.action });
    }
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

// ── AUTH ──────────────────────────────────────────────────────────────────────
function _login(idToken) {
  if (!idToken) return { ok: false, error: 'No Google token' };
  var resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return { ok: false, error: 'Invalid Google token' };
  var info = JSON.parse(resp.getContentText());
  if (info.aud !== CLIENT_ID) return { ok: false, error: 'Token was not issued for this app' };
  var verified = (info.email_verified === true || info.email_verified === 'true');
  var email = String(info.email || '').toLowerCase();
  if (!verified || !email) return { ok: false, error: 'Google email not verified' };
  if (_whitelist().indexOf(email) === -1) return { ok: false, error: 'Account ' + email + ' is not whitelisted for this app.' };
  var exp = Date.now() + SESSION_DAYS * 86400000;
  return { ok: true, email: email, expiresAt: exp, session: _signSession(email, exp) };
}

function _whitelist() { return WHITELIST.map(function (x) { return String(x).toLowerCase(); }); }

function _signSession(email, exp) {
  var payload = email + '|' + exp;
  var sig = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, SECRET));
  return Utilities.base64EncodeWebSafe(payload) + '.' + sig;
}

function _checkSession(session) {
  try {
    if (!session) return { ok: false };
    var parts = session.split('.');
    var payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
    var expect = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, SECRET));
    if (expect !== parts[1]) return { ok: false };
    var bits = payload.split('|'), email = bits[0], exp = Number(bits[1]);
    if (!exp || Date.now() > exp) return { ok: false };
    if (_whitelist().indexOf(email.toLowerCase()) === -1) return { ok: false };
    return { ok: true, email: email };
  } catch (err) { return { ok: false }; }
}

// Simple health check when the URL is opened in a browser.
function doGet() { return _json({ ok: true, service: 'jewel-rent-gateway' }); }

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function _sheet(spreadsheetId, name) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function _headers(sh) {
  var lastCol = sh.getLastColumn();
  if (lastCol === 0) return [];
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
}

// Return all rows as arrays, header row included (index 0). App maps by name.
function _read(spreadsheetId, name) {
  var sh = _sheet(spreadsheetId, name);
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return [];
  return sh.getRange(1, 1, lastRow, lastCol).getValues();
}

// Read MANY tabs in one call (one spreadsheet open, one HTTP round-trip) —
// much faster than one request per tab. Missing tabs come back as [].
function _readAll(spreadsheetId, names) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var out = {};
  (names || []).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { out[name] = []; return; }
    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    out[name] = (lastRow === 0 || lastCol === 0) ? [] : sh.getRange(1, 1, lastRow, lastCol).getValues();
  });
  return out;
}

// Append a row from a {header: value} object, ordered to the sheet's real header.
function _append(spreadsheetId, name, valuesObj) {
  var sh = _sheet(spreadsheetId, name);
  var headers = _headers(sh);
  if (headers.length === 0) {
    // No header yet — lay one down from the object's keys.
    headers = Object.keys(valuesObj);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  var row = headers.map(function (h) { return valuesObj.hasOwnProperty(h) ? valuesObj[h] : ''; });
  sh.appendRow(row);
  return sh.getLastRow();
}

// Update one row (1-based sheet row) from a {header: value} object, by name.
function _update(spreadsheetId, name, rowNum, valuesObj) {
  var sh = _sheet(spreadsheetId, name);
  var headers = _headers(sh);
  var current = sh.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  var row = headers.map(function (h, i) { return valuesObj.hasOwnProperty(h) ? valuesObj[h] : current[i]; });
  sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
  return true;
}

// Delete rows: either a single 1-based row number, or every row whose `field`
// column equals `value` (used to remove all legs of one transaction by Txn ID).
function _delete(spreadsheetId, name, rowNum, field, value) {
  var sh = _sheet(spreadsheetId, name);
  if (rowNum) { sh.deleteRow(rowNum); return 1; }
  var headers = _headers(sh);
  var col = headers.indexOf(field);
  if (col === -1) return 0;
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var data = sh.getRange(2, 1, last - 1, headers.length).getValues();
  var rows = [];
  for (var i = 0; i < data.length; i++) if (String(data[i][col]) === String(value)) rows.push(i + 2);
  for (var k = rows.length - 1; k >= 0; k--) sh.deleteRow(rows[k]); // bottom-up
  return rows.length;
}

// Ensure each tab exists and has every expected header; append missing headers
// at the end. Never inserts mid-sheet, never edits data rows.
function _ensureSchema(spreadsheetId, schema) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var changed = [];
  Object.keys(schema).forEach(function (name) {
    var want = schema[name];
    var sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, want.length).setValues([want]); changed.push(name + ':new'); return; }
    var have = _headers(sh);
    var missing = want.filter(function (h) { return have.indexOf(h) === -1; });
    if (missing.length) {
      sh.getRange(1, have.length + 1, 1, missing.length).setValues([missing]);
      changed.push(name + ':+' + missing.join(','));
    }
  });
  return changed;
}
