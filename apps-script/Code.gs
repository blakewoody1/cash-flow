/**
 * Cash Flow — Apps Script for your "Cash Flow" Google Sheet
 *
 * BACKUP COPY. The live copy is in the Cash Flow sheet → Extensions → Apps Script.
 * The KEY below is a placeholder on purpose — never commit your real key to this public repo.
 *
 * This one script does two jobs:
 *   1. Takes entries from your iPhone Shortcuts (Log Expense, Log Savings, Sync Cash Flow).
 *   2. Lets the Cash Flow web app read and update the sheet.
 *
 * Both use the same password (KEY below). After changing anything here:
 *   Save → Deploy → Manage deployments → ✎ (pencil) → Version: New version → Deploy.
 *   (That keeps the same web app URL, so nothing else needs a new address.)
 *
 * Any request with a wrong key is written to an "Access log" tab in the Cash Flow sheet.
 */
const KEY = 'CHANGE-ME-to-a-long-random-password';

// The only sheets the app may open (sheet ID → account name). Add a line here for a new account.
const SHEETS = {
  '14yvG4FmByEp2Te4gIPQ5FPQ2n6hwSN-quB7b95WRQ1U': 'Cash Flow',
  '109ZOrysstYv1adeMnFclmQD_dmK92RVXOx0kcpql18E': 'WNC'
};

function sheetFor_(id) {
  if (!id) return SpreadsheetApp.getActive();
  if (!SHEETS[id]) throw new Error('That sheet is not in the SHEETS list in the Apps Script.');
  return SpreadsheetApp.openById(id);
}

// Wrong-key attempts go to an "Access log" tab in the Cash Flow sheet (created the first time)
function logBad_(source, detail) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('Access log');
    if (!sh) { sh = ss.insertSheet('Access log'); sh.appendRow(['Time', 'Source', 'What was tried', 'Result']); sh.setFrozenRows(1); }
    sh.appendRow([new Date(), source, String(detail || '').slice(0, 60), 'Refused: wrong key']);
  } catch (err) { /* never let logging break a request */ }
}

/* ---------- entry points ---------- */
function doGet() {
  return ContentService.createTextOutput('Cash Flow backend is running.');
}

function doPost(e) {
  const body = (e && e.postData && e.postData.contents) || '';
  let req = null;
  try { req = JSON.parse(body); } catch (err) { /* not one JSON object: it's from the Shortcuts */ }
  if (req && req.tool) return json_(api_(req));
  return shortcutLog_(body);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 1. Shortcuts: one JSON entry per line ---------- */
function shortcutLog_(contents) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const out = [], bad = [];
    const byName = name => { const id = Object.keys(SHEETS).filter(k => SHEETS[k].toLowerCase() === String(name).toLowerCase())[0]; return id ? SpreadsheetApp.openById(id) : null; };
    const text = contents.replace(/[“”„″]/g, '"').replace(/[‘’]/g, "'");
    text.split('\n').map(s => s.trim()).filter(Boolean).forEach(line => {
      let d;
      try { d = JSON.parse(line); } catch (err) { bad.push(line.slice(0, 40)); return; }
      if (d.key !== KEY) { bad.push('wrong key'); logBad_('Shortcut', (d.type || '') + ' ' + (d.what || '')); return; }
      const ss = (d.account && byName(d.account)) || SpreadsheetApp.getActive();
      const sh = ss.getSheetByName(d.type === 'savings' ? 'Savings Log' : 'Spending Log');
      if (d.id) {
        const n = sh.getLastRow();
        const ids = n > 1 ? sh.getRange(2, 4, n - 1, 1).getValues().flat().map(String) : [];
        if (ids.includes(String(d.id))) return;
      }
      const now = new Date();
      let day = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      const p = String(d.date || '').split('/');
      if (p.length === 3) day = new Date(+p[2], +p[0] - 1, +p[1]);
      sh.appendRow([day, d.what, Number(d.amount), d.id || '']);
      out.push(d.what + ' $' + d.amount);
    });
    let msg = out.length ? 'Logged ' + out.join(', ') : (bad.length ? "Couldn't read" : 'Nothing new');
    if (bad.length) msg += ' (skipped: ' + bad.join(' | ') + ')';
    return ContentService.createTextOutput(msg);
  } finally { lock.releaseLock(); }
}

/* ---------- 2. Web app: small read/write API ---------- */
function api_(req) {
  if (req.key !== KEY) { logBad_('App', req.tool); return { error: 'Wrong key', code: 'needs_reauth' }; }
  const write = req.tool === 'update_values' || req.tool === 'insert_dimension';
  const lock = write ? LockService.getScriptLock() : null;
  if (lock) lock.waitLock(30000);
  try {
    const i = req.input || {}, ss = sheetFor_(i.spreadsheetId);
    switch (req.tool) {
      case 'get_spreadsheet': return { result: getSpreadsheet_(ss, i) };
      case 'get_values': return { result: getValues_(ss, i) };
      case 'update_values': return { result: updateValues_(ss, i) };
      case 'insert_dimension': return { result: insertDimension_(ss, i) };
      default: return { error: 'Unknown request: ' + req.tool, code: 'bad_request' };
    }
  } catch (err) {
    return { error: String((err && err.message) || err), code: 'tool_error' };
  } finally {
    if (lock) lock.releaseLock();
  }
}

// Google Sheets date number (days since Dec 30, 1899) in the sheet's own time zone
function serial_(d, tz) {
  const s = Utilities.formatDate(d, tz, 'yyyy-MM-dd').split('-');
  return Math.round((Date.UTC(+s[0], +s[1] - 1, +s[2]) - Date.UTC(1899, 11, 30)) / 864e5);
}

// Only the part of a range that actually has data (saves time and data)
function usedPart_(rg) {
  const sh = rg.getSheet();
  const rows = Math.min(rg.getNumRows(), sh.getLastRow() - rg.getRow() + 1);
  return rows > 0 ? sh.getRange(rg.getRow(), rg.getColumn(), rows, rg.getNumColumns()) : null;
}

function getSpreadsheet_(ss, i) {
  if (!i.includeGridData) {
    return {
      properties: { title: ss.getName(), spreadsheetId: ss.getId() },
      sheets: ss.getSheets().map(s => ({ properties: { title: s.getName(), sheetId: s.getSheetId() } }))
    };
  }
  const tz = ss.getSpreadsheetTimeZone();
  return {
    sheets: (i.ranges || []).map(a1 => {
      const rg = ss.getRange(a1), sh = rg.getSheet(), part = usedPart_(rg), rowData = [];
      if (part) {
        const f = part.getFormulas(), v = part.getValues();
        for (let r = 0; r < v.length; r++) {
          const values = [];
          for (let c = 0; c < v[r].length; c++) {
            const val = v[r][c], form = f[r][c], cell = {};
            if (form) cell.userEnteredValue = { formulaValue: form };
            else if (val instanceof Date) cell.userEnteredValue = { numberValue: serial_(val, tz) };
            else if (typeof val === 'number') cell.userEnteredValue = { numberValue: val };
            else if (typeof val === 'boolean') cell.userEnteredValue = { stringValue: String(val).toUpperCase() };
            else if (val !== '') cell.userEnteredValue = { stringValue: String(val) };
            if (val instanceof Date) cell.effectiveValue = { numberValue: serial_(val, tz) };
            else if (typeof val === 'number') cell.effectiveValue = { numberValue: val };
            else if (val !== '' && val != null) cell.effectiveValue = { stringValue: String(val) };
            values.push(cell);
          }
          while (values.length && !values[values.length - 1].userEnteredValue) values.pop();
          rowData.push({ values: values });
        }
        while (rowData.length && !rowData[rowData.length - 1].values.length) rowData.pop();
      }
      return { properties: { title: sh.getName(), sheetId: sh.getSheetId() }, data: [{ startRow: rg.getRow() - 1, rowData: rowData }] };
    })
  };
}

function getValues_(ss, i) {
  const part = usedPart_(ss.getRange(i.range));
  const values = part ? part.getDisplayValues().map(row => {
    const r = row.slice(); while (r.length && r[r.length - 1] === '') r.pop(); return r;
  }) : [];
  while (values.length && !values[values.length - 1].length) values.pop();
  return { range: i.range, values: values };
}

// Values are entered as if typed: "=..." becomes a formula, "11/2/2026" becomes a date
function updateValues_(ss, i) {
  const rg = ss.getRange(i.range), vals = i.values || [];
  if (!vals.length) return { status: 'success' };
  const width = Math.max.apply(null, vals.map(r => r.length));
  const tz = ss.getSpreadsheetTimeZone();
  const grid = vals.map(r => {
    const x = r.map(v => {
      if (v == null) return '';
      // the app sends dates as M/D/YYYY — store them as real dates in the sheet's time zone
      if (typeof v === 'string' && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(v)) return Utilities.parseDate(v, tz, 'M/d/yyyy');
      return v;
    });
    while (x.length < width) x.push('');
    return x;
  });
  rg.getSheet().getRange(rg.getRow(), rg.getColumn(), grid.length, width).setValues(grid);
  SpreadsheetApp.flush();
  return { updatedRange: i.range, status: 'success' };
}

function insertDimension_(ss, i) {
  const sh = ss.getSheets().filter(s => s.getSheetId() === i.sheetId)[0];
  if (!sh) throw new Error('No tab with id ' + i.sheetId);
  const n = (i.endIndex || 0) - (i.startIndex || 0);
  if (n <= 0) return { status: 'success' };
  if (i.dimension === 'COLUMNS') {
    if (i.startIndex > 0 && i.inheritFromBefore !== false) sh.insertColumnsAfter(i.startIndex, n); else sh.insertColumnsBefore(i.startIndex + 1, n);
  } else {
    if (i.startIndex > 0 && i.inheritFromBefore !== false) sh.insertRowsAfter(i.startIndex, n); else sh.insertRowsBefore(i.startIndex + 1, n);
  }
  return { status: 'success' };
}
