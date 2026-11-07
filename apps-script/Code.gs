/**
 * BACKUP COPY. The live copy is in the Cash Flow sheet → Extensions → Apps Script.
 * The KEY below is a placeholder on purpose — never commit your real key to this public repo.
 *
 * Cash Flow — Apps Script for your "Cash Flow" Google Sheet
 *
 * Lets the Cash Flow web app read and update your sheets, your Bills calendar and your Affirm tracker.
 * (The Log Expense / Log Savings Shortcuts were retired Nov 7, 2026: everything goes through the app.)
 *
 * The app sends the password (KEY below) with every request. After changing anything here:
 *   Save → Deploy → Manage deployments → ✎ (pencil) → Version: New version → Deploy.
 *   (That keeps the same web app URL, so nothing else needs a new address.)
 *
 * Any request with a wrong key is written to an "Access log" tab in the Cash Flow sheet.
 *
 * Safety rules for the app's requests (so a leaked key can't do much):
 *   - it can only write to month tabs ("Nov 2026") and the Spending / Savings Log, columns A–I, a few rows at a time
 *   - formulas are only the running-balance and total formulas the app uses; anything else is saved as plain text
 *   - new rows are added in one step here, and each change has an ID so a retry never adds it twice
 *   - calendar changes only touch the Bills calendar; Affirm changes only update a plan's balance and add to Payment History
 *   - after 30 wrong keys in 10 minutes it stops answering for a while, and the Access log keeps only the last 500 lines
 *
 * Keeping the sheets current (runs by itself once set up):
 *   - On the 1st of every month it adds month tabs so there are always about 14 months ahead,
 *     copied from your latest month: monthly bills on the same day, weekly bills and paychecks on the
 *     same weekday, plus the weekly Everyday spending / Moved to savings rows.
 *   - Each January it moves month tabs and log entries older than last year into a separate
 *     "<name> Archive" spreadsheet in your Drive, so the sheet (and the app) stay fast. Nothing is deleted.
 *   To turn this on, run setupSchedule once from the editor (pick it in the function menu, then ▷ Run).
 */
const KEY = 'CHANGE-ME-to-a-long-random-password';

// The only sheets the app may open (sheet ID → account name). Add a line here for a new account.
const SHEETS = {
  '14yvG4FmByEp2Te4gIPQ5FPQ2n6hwSN-quB7b95WRQ1U': 'Cash Flow',
  '109ZOrysstYv1adeMnFclmQD_dmK92RVXOx0kcpql18E': 'WNC'
};

// Bills calendar (Google Calendar) and the Affirm Payment Plans Tracker sheet
const BILLS_CAL = '8f3aaa334c8bf14773c60647529fdf1dba99beedcdaf2cce6ff0a94646347551@group.calendar.google.com';
const AFFIRM = '1U3zL_us7_RUKxhL_WSevIIYmKSWRL8p02rCdlJGJTnw';

const MONTHS_AHEAD = 14;    // keep month tabs this many months past the current month
const KEEP_YEARS_BACK = 1;  // archive month tabs and log entries from before January of last year

const opened_ = {};
function sheetFor_(id) {
  const active = SpreadsheetApp.getActive();
  if (!id || (active && id === active.getId())) return active;   // the sheet this script lives in opens fastest
  if (!SHEETS[id]) throw new Error('That sheet is not in the SHEETS list in the Apps Script.');
  return opened_[id] || (opened_[id] = SpreadsheetApp.openById(id));
}

// Wrong-key attempts go to an "Access log" tab in the Cash Flow sheet (created the first time)
function logBad_(source, detail) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('Access log');
    if (!sh) { sh = ss.insertSheet('Access log'); sh.appendRow(['Time', 'Source', 'What was tried', 'Result']); sh.setFrozenRows(1); }
    if (sh.getLastRow() > 500) sh.deleteRows(2, sh.getLastRow() - 500);
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
  try { req = JSON.parse(body); } catch (err) { /* not JSON */ }
  if (req && req.tool) return json_(api_(req));
  return json_({ error: 'Not a Cash Flow app request.' });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 2. Web app: small read/write API ---------- */
function api_(req) {
  const cache = CacheService.getScriptCache(), tries = +(cache.get('bad') || 0);
  if (tries >= 30) return { error: 'Too many wrong keys. Try again in 10 minutes.', code: 'needs_reauth' };
  if (req.key !== KEY) { cache.put('bad', String(tries + 1), 600); logBad_('App', req.tool); return { error: 'Wrong key', code: 'needs_reauth' }; }
  const write = WRITES_.indexOf(req.tool) >= 0;
  const lock = write ? LockService.getScriptLock() : null;
  if (lock) lock.waitLock(30000);
  try {
    // several reads in one trip (the app uses this to open faster on a slow signal)
    if (req.tool === 'batch') return { result: { results: ((req.input || {}).calls || []).map(c => run_(c.tool, c.input || {}, true)) } };
    return run_(req.tool, req.input || {}, false);
  } finally {
    if (lock) lock.releaseLock();
  }
}
const WRITES_ = ['update_values', 'insert_dimension', 'add_item', 'add_log', 'cal_add', 'cal_update', 'cal_delete', 'affirm_pay'];
function run_(tool, i, readOnly) {
  try {
    if (/^cal_/.test(tool)) {
      if (tool === 'cal_list') return { result: calList_(i) };
      if (readOnly) throw new Error('Read-only request.');
      if (tool === 'cal_add') return { result: calAdd_(i) };
      if (tool === 'cal_update') return { result: calUpdate_(i) };
      if (tool === 'cal_delete') return { result: calDelete_(i) };
    }
    if (tool === 'affirm_get') return { result: affirmGet_() };
    if (tool === 'affirm_pay') { if (readOnly) throw new Error('Read-only request.'); return { result: affirmPay_(i) }; }
    const ss = sheetFor_(i.spreadsheetId);
    switch (tool) {
      case 'get_spreadsheet': return { result: getSpreadsheet_(ss, i) };
      case 'get_values': return { result: getValues_(ss, i) };
      case 'update_values': if (!readOnly) return { result: updateValues_(ss, i) }; break;
      case 'insert_dimension': if (!readOnly) return { result: insertDimension_(ss, i) }; break;
      case 'add_item': if (!readOnly) return { result: addItem_(ss, i) }; break;
      case 'add_log': if (!readOnly) return { result: addLog_(ss, i) }; break;
    }
    return { error: 'Unknown request: ' + tool, code: 'bad_request' };
  } catch (err) {
    return { error: String((err && err.message) || err), code: 'tool_error' };
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
      properties: { title: ss.getName(), spreadsheetId: ss.getId(), features: ['batch', 'add_item', 'add_log', 'guard', 'calendar', 'affirm'] },
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
  checkWrite_(rg.getSheet(), rg.getColumn(), rg.getColumn() + width - 1, vals.length);
  const tz = ss.getSpreadsheetTimeZone();
  const grid = vals.map(r => {
    const x = r.map(v => {
      if (v == null) return '';
      if (typeof v === 'string') v = safeText_(v);
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
  if (i.dimension === 'COLUMNS') throw new Error('The app may only add rows.');
  checkWrite_(sh, 1, 1, n);
  if (isMonth_(sh) && i.startIndex < 12) throw new Error('Rows can only be added below the header.');
  if (i.dimension === 'COLUMNS') {
    if (i.startIndex > 0 && i.inheritFromBefore !== false) sh.insertColumnsAfter(i.startIndex, n); else sh.insertColumnsBefore(i.startIndex + 1, n);
  } else {
    if (i.startIndex > 0 && i.inheritFromBefore !== false) sh.insertRowsAfter(i.startIndex, n); else sh.insertRowsBefore(i.startIndex + 1, n);
  }
  return { status: 'success' };
}

/* ---------- 2b. Safety rules and one-step adds ---------- */
const LOGS_ = ['Spending Log', 'Savings Log'];
const MONTH_RE_ = /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{4}$/;
function isMonth_(sh) { return MONTH_RE_.test(sh.getName()); }
// The app only ever writes to month tabs and the two logs, a handful of cells at a time
function checkWrite_(sh, c1, c2, rows) {
  const name = sh.getName(), log = LOGS_.indexOf(name) >= 0;
  if (!log && !isMonth_(sh)) throw new Error('The app can only change month tabs and the logs, not “' + name + '”.');
  if (c1 < 1 || c2 > (log ? 5 : 9)) throw new Error('That change is outside the columns the app uses.');
  if (rows > 20) throw new Error('That change is bigger than the app ever makes.');
}
// Only the formulas the app writes (running balance and the month totals). Anything else starting with = is saved as plain text.
const OK_FORMULAS_ = [/^=F\d+\+D\d+-E\d+$/, /^=SUM\(D\d+:D\d+\)$/, /^=SUM\(E\d+:E\d+\)-E6-E9$/, /^=MIN\(F\d+:F\d+\)$/, /^=F\d+$/, /^=B4$/];
function safeText_(v) {
  if (v.charAt(0) !== '=') return v;
  const f = v.replace(/\s+/g, '');
  return OK_FORMULAS_.some(re => re.test(f)) ? f : "'" + v;
}

// Each change from the app carries an ID; remember the ones already done (kept about 90 days)
function opKey_(ss, id) {
  return 'op:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, ss.getId() + '|' + id));
}
function opDone_(ss, id) { return !!(id && PropertiesService.getScriptProperties().getProperty(opKey_(ss, id))); }
function markOp_(ss, id) { if (id) PropertiesService.getScriptProperties().setProperty(opKey_(ss, id), String(Date.now())); }
function pruneOps_() {
  const props = PropertiesService.getScriptProperties(), all = props.getProperties(), old = Date.now() - 90 * 864e5;
  Object.keys(all).forEach(k => { if (k.indexOf('op:') === 0 && +all[k] < old) props.deleteProperty(k); });
}

// Adds one bill or income to a month tab in a single step: finds its spot by date, adds the row, and keeps the running balance and totals linked
function addItem_(ss, i) {
  if (!i.opId) throw new Error('Missing change ID.');
  if (opDone_(ss, i.opId)) return { status: 'success', duplicate: true };
  const sh = ss.getSheetByName(String(i.title || ''));
  if (!sh || !isMonth_(sh)) throw new Error('No month tab called ' + i.title);
  const tz = ss.getSpreadsheetTimeZone();
  if (!/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(String(i.date))) throw new Error('Bad date.');
  const date = Utilities.parseDate(i.date, tz, 'M/d/yyyy'), target = serial_(date, tz);
  const amt = Number(i.amt);
  if (!isFinite(amt) || Math.abs(amt) > 1e7) throw new Error('Bad amount.');
  if (i.kind !== 'income' && i.kind !== 'bill') throw new Error('Bad kind.');

  const last = sh.getLastRow(), vals = sh.getRange(1, 1, last, 9).getValues(), forms = sh.getRange(1, 1, last, 9).getFormulas();
  let hdr = 0;
  for (let r = 0; r < last; r++) if (vals[r][0] === 'Date' && /^Running/i.test(String(vals[r][5]))) { hdr = r + 1; break; }
  if (!hdr) throw new Error('No Date / Running Balance header on ' + i.title);
  const first = hdr + 1;
  let end = first - 1;
  const real = [];
  for (let r = first; r <= last; r++) {
    const v = vals[r - 1], f = forms[r - 1];
    if (!f[5]) break;
    end = r;
    const blank = !String(v[2] || '').trim() && v[3] === '' && v[4] === '' && !f[3] && !f[4];
    if (!blank && v[0] instanceof Date) real.push({ r: r, serial: serial_(v[0], tz) });
  }
  if (end < first) throw new Error('No rows on ' + i.title);
  if (real.length && target < real[0].serial) throw new Error('“' + i.title + '” starts later than that date.');
  let pos = first;
  real.forEach(x => { if (x.serial <= target) pos = x.r; });
  const n = pos + 1, wd = DOW_[new Date(Date.UTC(1899, 11, 30) + target * 864e5).getUTCDay()];

  sh.insertRowsAfter(pos, 1);
  sh.getRange(n, 1, 1, 9).setValues([[date, wd, safeText_(String(i.desc || '').slice(0, 200)),
    i.kind === 'income' ? amt : '', i.kind === 'bill' ? amt : '', '=F' + pos + '+D' + n + '-E' + n,
    safeText_(String(i.note || '').slice(0, 300)), '', i.cleared ? '✓' : '']]);
  if (i.cleared && !vals[hdr - 1][8]) sh.getRange(hdr, 9).setValue('Cleared');
  if (pos < end) sh.getRange(n + 1, 6).setFormula('=F' + n + '+D' + (n + 1) + '-E' + (n + 1));
  else {
    const L = end + 1;
    sh.getRange(4, 5, 2, 1).setFormulas([['=SUM(D' + first + ':D' + L + ')'], ['=SUM(E' + first + ':E' + L + ')-E6-E9']]);
    sh.getRange(7, 5, 2, 1).setFormulas([['=MIN(F' + first + ':F' + L + ')'], ['=F' + L]]);
  }
  SpreadsheetApp.flush();
  markOp_(ss, i.opId);
  return { status: 'success', row: n };
}

// Adds one entry to the Spending or Savings Log; the change ID goes in column D (same place the Shortcuts put theirs) so it's never added twice
function addLog_(ss, i) {
  if (!i.opId) throw new Error('Missing change ID.');
  if (LOGS_.indexOf(i.tab) < 0) throw new Error('Unknown log ' + i.tab);
  const sh = ss.getSheetByName(i.tab), id = 'app:' + i.opId;
  const n = sh.getLastRow();
  if (n > 1 && sh.getRange(2, 4, n - 1, 1).getValues().some(r => String(r[0]) === id)) return { status: 'success', duplicate: true };
  const tz = ss.getSpreadsheetTimeZone();
  if (!/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(String(i.date))) throw new Error('Bad date.');
  const amt = Number(i.amt);
  if (!isFinite(amt) || Math.abs(amt) > 1e7) throw new Error('Bad amount.');
  const r = n + 1;
  sh.getRange(r, 1, 1, 5).setValues([[Utilities.parseDate(i.date, tz, 'M/d/yyyy'), safeText_(String(i.text || '').slice(0, 200)), amt, id, i.cleared ? '✓' : '']]);
  if (i.cleared && !sh.getRange(1, 5).getValue()) sh.getRange(1, 5).setValue('Cleared');
  SpreadsheetApp.flush();
  return { status: 'success', row: r };
}

/* ---------- 2c. Bills calendar ---------- */
// Run this once from the editor (pick "authorize" in the function menu, then ▷ Run) so Google asks
// for permission to use your Bills calendar and the Affirm tracker. It changes nothing.
function authorize() {
  const cal = billsCal_(), plans = affirmSheets_().plans;
  Logger.log('Bills calendar: ' + cal.getName() + ' · Affirm plans tab rows: ' + plans.getLastRow());
}

function billsCal_() {
  const c = CalendarApp.getCalendarById(BILLS_CAL);
  if (!c) throw new Error("Can't open the Bills calendar.");
  return c;
}
// Some sheets report no usable time zone, so fall back to the script's (or Eastern)
function isoDate_(d, tz) {
  const zones = [tz, Session.getScriptTimeZone(), 'America/New_York'];
  for (let k = 0; k < zones.length; k++) {
    if (typeof zones[k] !== 'string' || !zones[k]) continue;
    try { return Utilities.formatDate(d, zones[k], 'yyyy-MM-dd'); } catch (err) { /* try the next */ }
  }
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}
function parseYmd_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s))) throw new Error('Bad date.');
  const p = s.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]);
}
// Events between two dates (yyyy-mm-dd), at most about 3 months at a time
function calList_(i) {
  const cal = billsCal_(), tz = Session.getScriptTimeZone();
  const from = parseYmd_(i.from), to = parseYmd_(i.to);
  if (to - from > 100 * 864e5 || to < from) throw new Error('Pick up to about 3 months.');
  const end = new Date(to.getTime() + 864e5);
  return { events: cal.getEvents(from, end).map(ev => ({
    id: ev.getId(), title: ev.getTitle(), date: isoDate_(ev.isAllDayEvent() ? ev.getAllDayStartDate() : ev.getStartTime(), tz),
    allDay: ev.isAllDayEvent(), note: ev.getDescription() || '', repeats: ev.isRecurringEvent()
  })) };
}
// The calendar event matching an ID on a given day (one day of a repeating event)
function findEvent_(id, date) {
  const day = parseYmd_(date), evs = billsCal_().getEvents(day, new Date(day.getTime() + 864e5)).filter(ev => ev.getId() === id);
  if (!evs.length) throw new Error("That calendar event isn't there anymore.");
  return evs[0];
}
const clip_ = (v, n) => String(v == null ? '' : v).slice(0, n);
function calAdd_(i) {
  const cal = billsCal_(), title = clip_(i.title, 120).trim(), day = parseYmd_(i.date), note = clip_(i.note, 500);
  if (!title) throw new Error('Give the event a name.');
  if (i.opId) {   // already added on an earlier try?
    const same = cal.getEvents(day, new Date(day.getTime() + 864e5)).filter(ev => ev.getTag('opId') === String(i.opId));
    if (same.length) return { status: 'success', duplicate: true, id: same[0].getId() };
  }
  let ev;
  if (i.repeat && i.repeat !== 'never') {
    const r = CalendarApp.newRecurrence();
    const rule = i.repeat === 'weekly' ? r.addWeeklyRule() : i.repeat === 'biweekly' ? r.addWeeklyRule().interval(2) : r.addMonthlyRule();
    if (i.until) rule.until(parseYmd_(i.until));
    ev = cal.createAllDayEventSeries(title, day, r, { description: note });
  } else ev = cal.createAllDayEvent(title, day, { description: note });
  if (i.opId) ev.setTag('opId', String(i.opId));
  return { status: 'success', id: ev.getId() };
}
// scope 'one' = just that day; 'all' = the whole repeating series
function calUpdate_(i) {
  const ev = findEvent_(i.id, i.date), title = clip_(i.title, 120).trim(), note = clip_(i.note, 500);
  if (!title) throw new Error('Give the event a name.');
  if (i.scope === 'all' && ev.isRecurringEvent()) {
    const series = ev.getEventSeries(); series.setTitle(title); series.setDescription(note);
  } else {
    ev.setTitle(title); ev.setDescription(note);
    if (i.newDate && i.newDate !== i.date) ev.setAllDayDate(parseYmd_(i.newDate));
  }
  return { status: 'success' };
}
function calDelete_(i) {
  const ev = findEvent_(i.id, i.date);
  if (i.scope === 'all' && ev.isRecurringEvent()) ev.getEventSeries().deleteEventSeries(); else ev.deleteEvent();
  return { status: 'success' };
}

/* ---------- 2d. Affirm Payment Plans Tracker ---------- */
// Enter Plans: A Remaining Balance, B Plan / Merchant, C Next Due Date, D Payments Left, E APR, F Monthly Payment, G Plan ID, H Status, I Paid Off Date
// Payment History: Payment ID, Plan ID, Merchant, Payment Date, Amount Paid, Balance Before, Balance After, Status, Note, Recorded At
function affirmSheets_() {
  const ss = SpreadsheetApp.openById(AFFIRM);
  const plans = ss.getSheetByName('Enter Plans'), hist = ss.getSheetByName('Payment History');
  if (!plans || !hist) throw new Error('The Affirm tracker needs "Enter Plans" and "Payment History" tabs.');
  let tz = ''; try { tz = ss.getSpreadsheetTimeZone(); } catch (err) { /* fall back below */ }
  return { ss: ss, plans: plans, hist: hist, tz: String(tz || Session.getScriptTimeZone() || 'America/New_York') };
}
function affirmGet_() {
  const a = affirmSheets_(), n = a.plans.getLastRow(), h = a.hist.getLastRow();
  const rows = n > 1 ? a.plans.getRange(2, 1, n - 1, 9).getValues() : [];
  const hist = h > 1 ? a.hist.getRange(Math.max(2, h - 49), 1, Math.min(50, h - 1), 10).getValues() : [];
  const d = v => v instanceof Date ? isoDate_(v, a.tz) : String(v || '');
  return {
    plans: rows.map((r, k) => ({ row: k + 2, balance: +r[0] || 0, name: String(r[1] || ''), due: d(r[2]), left: +r[3] || 0,
      apr: (v => v > 1 ? v / 100 : v)(typeof r[4] === 'number' ? r[4] : parseFloat(String(r[4])) / 100 || 0), payment: +r[5] || 0, id: String(r[6] || ''), status: String(r[7] || ''), paidOff: d(r[8]) }))
      .filter(p => p.name || p.balance),
    history: hist.map(r => ({ id: String(r[0]), plan: String(r[1]), name: String(r[2]), date: d(r[3]), amount: +r[4] || 0, before: +r[5] || 0, after: +r[6] || 0, status: String(r[7]), note: String(r[8]) }))
  };
}
// Records a payment toward one plan: lowers its balance, adds a Payment History line, marks it paid off at $0
function affirmPay_(i) {
  if (!i.opId) throw new Error('Missing change ID.');
  const a = affirmSheets_(), amt = Math.round(Number(i.amount) * 100) / 100;
  if (!isFinite(amt) || amt <= 0 || amt > 1e5) throw new Error('Bad amount.');
  const h = a.hist.getLastRow(), pid = 'app:' + i.opId;
  if (h > 1 && a.hist.getRange(2, 1, h - 1, 1).getValues().some(r => String(r[0]) === pid)) return { status: 'success', duplicate: true };
  const n = a.plans.getLastRow(), rows = n > 1 ? a.plans.getRange(2, 1, n - 1, 9).getValues() : [];
  const k = rows.findIndex(r => String(r[6]) === String(i.planId));
  if (k < 0) throw new Error("Couldn't find that plan in the tracker.");
  const r = k + 2, before = +rows[k][0] || 0, after = Math.max(0, Math.round((before - amt) * 100) / 100), pay = +rows[k][5] || 0;
  if (amt > before + 0.005) throw new Error('That is more than the ' + rows[k][1] + ' balance ($' + before.toFixed(2) + ').');
  const date = /^\d{4}-\d{2}-\d{2}$/.test(String(i.date)) ? parseYmd_(i.date) : new Date();
  a.plans.getRange(r, 1).setValue(after);
  if (after <= 0) { a.plans.getRange(r, 4).setValue(0); a.plans.getRange(r, 8).setValue('PAID OFF'); a.plans.getRange(r, 9).setValue(date); }
  else if (pay > 0) a.plans.getRange(r, 4).setValue(Math.ceil(after / pay - 1e-9));
  a.hist.appendRow([pid, String(rows[k][6]), safeText_(String(rows[k][1])), date, amt, before, after, after <= 0 ? 'PAID OFF' : (i.kind === 'regular' ? 'PAYMENT' : 'EXTRA'), safeText_(clip_(i.note, 200)), new Date()]);
  SpreadsheetApp.flush();
  return { status: 'success', before: before, after: after, paidOff: after <= 0 };
}

/* ---------- 3. Keeping the sheets current ---------- */
const MON_ = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const MONTH_NAME_ = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const DOW_ = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

// Run this once from the editor: it schedules the monthly job and runs it right away.
function setupSchedule() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'monthlyMaintenance').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('monthlyMaintenance').timeBased().onMonthDay(1).atHour(3).create();
  const report = monthlyMaintenance();
  console.log('Scheduled: runs on the 1st of every month around 3 AM.\n' + report);
}

function monthlyMaintenance() {
  const lock = LockService.getScriptLock(); lock.waitLock(60000);
  const lines = [];
  try {
    Object.keys(SHEETS).forEach(id => {
      const name = SHEETS[id];
      try {
        const ss = sheetFor_(id);
        const added = addMonths_(ss), archived = archiveOld_(ss, name);
        lines.push(name + ': ' + ([added.length ? 'added ' + added.join(', ') : '', archived.length ? 'archived ' + archived.join(', ') : ''].filter(Boolean).join('; ') || 'already up to date'));
      } catch (err) { lines.push(name + ': problem — ' + ((err && err.message) || err)); }
    });
    try { pruneOps_(); } catch (err) { /* not important */ }
  } finally { lock.releaseLock(); }
  const report = lines.join('\n');
  console.log(report);
  return report;
}

function monthTabs_(ss) {
  return ss.getSheets().map(sh => {
    const m = sh.getName().match(/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) (\d{4})$/);
    return m ? { sh: sh, title: sh.getName(), y: +m[2], m: MON_.indexOf(m[1]), key: (+m[2]) * 12 + MON_.indexOf(m[1]) } : null;
  }).filter(Boolean).sort((a, b) => a.key - b.key);
}

function addMonths_(ss) {
  const tabs = monthTabs_(ss), made = [];
  if (!tabs.length) return made;
  const now = new Date(), target = now.getFullYear() * 12 + now.getMonth() + MONTHS_AHEAD;
  let guard = 0;
  while (tabs[tabs.length - 1].key < target && guard++ < 36) {
    const t = makeNextMonth_(ss, tabs[tabs.length - 1], tabs[tabs.length - 2]);
    tabs.push(t); made.push(t.title);
  }
  return made;
}

// day numbers (days since 1970-01-01) keep the date math free of time-zone surprises
function dayNum_(y, m, d) { return Math.round(Date.UTC(y, m, d) / 864e5); }
function ymd_(n) { const d = new Date(n * 864e5); return { y: d.getUTCFullYear(), m: d.getUTCMonth(), d: d.getUTCDate(), w: d.getUTCDay() }; }
function daysIn_(y, m) { return new Date(Date.UTC(y, m + 1, 0)).getUTCDate(); }

// Reads a month tab into plain rows
function readMonth_(t, tz) {
  const sh = t.sh, last = sh.getLastRow();
  const vals = sh.getRange(1, 1, last, 9).getValues(), forms = sh.getRange(1, 1, last, 9).getFormulas();
  let hdr = 0;
  for (let r = 0; r < last; r++) if (vals[r][0] === 'Date' && /^Running/i.test(String(vals[r][5]))) { hdr = r + 1; break; }
  if (!hdr) throw new Error('No Date / Running Balance header on ' + t.title);
  const rows = [];
  for (let r = hdr + 1; r <= last; r++) {
    const v = vals[r - 1], f = forms[r - 1];
    if (!f[5]) break;
    let day = null;
    if (v[0] instanceof Date) { const p = Utilities.formatDate(v[0], tz, 'yyyy-M-d').split('-'); day = dayNum_(+p[0], +p[1] - 1, +p[2]); }
    rows.push({ r: r, day: day, desc: String(v[2] || '').trim(), inc: typeof v[3] === 'number' ? v[3] : 0, bill: typeof v[4] === 'number' ? v[4] : 0,
      incF: f[3], billF: f[4], note: String(v[6] || ''), hF: f[7], first: r === hdr + 1 });
  }
  return { title: t.title, y: t.y, m: t.m, hdr: hdr, rows: rows, a1: String(vals[0][0] || ''), a2: String(vals[1][0] || '') };
}

function kind_(row) {
  if (row.first || /^starting balance/i.test(row.desc)) return 'start';
  if (row.billF && /Spending Log/.test(row.billF)) return 'spend';
  if (row.billF && /Savings Log/.test(row.billF)) return 'save';
  if (!row.desc && !row.inc && !row.bill) return 'blank';
  if (row.inc) return 'inc';
  return 'bill';
}

// Works out next month's rows from the latest month (and the one before it for the repeat spacing). No sheet access.
// A bill or paycheck whose note says "ends 3/23/2028" isn't copied into months after that date
function endsOn_(note) {
  const m = /\bends\s+(\d{1,2})\/(\d{1,2})\/(\d{4})/i.exec(String(note || ''));
  return m ? dayNum_(+m[3], +m[1] - 1, +m[2]) : null;
}
function planMonth_(last, prev, y, m) {
  const start = dayNum_(y, m, 1), end = dayNum_(y, m, daysIn_(y, m));
  const keyOf = r => r.desc + '|' + r.note + '|' + kind_(r);
  const groups = {}, order = [];
  last.rows.forEach((r, i) => {
    const k = kind_(r); if ((k !== 'inc' && k !== 'bill') || r.day == null) return;
    const key = keyOf(r);
    if (!groups[key]) { groups[key] = { kind: k, rows: [], prevDays: [], idx: i }; order.push(key); }
    groups[key].rows.push(r);
  });
  if (prev) prev.rows.forEach(r => { const g = groups[keyOf(r)]; if (g && r.day != null) g.prevDays.push(r.day); });

  const out = [];   // {day, rank, idx, kind, desc, inc, bill, note, src}
  order.forEach(key => {
    const g = groups[key], lastRow = g.rows[g.rows.length - 1];
    const days = g.prevDays.concat(g.rows.map(r => r.day)).sort((a, b) => a - b);
    const diffs = days.slice(1).map((d, i) => d - days[i]);
    const step = diffs.indexOf(7) >= 0 ? 7 : diffs.indexOf(14) >= 0 ? 14 : 0;
    const ends = endsOn_(lastRow.note);   // "ends 3/23/2028" in the note stops it after that date
    const add = (day, r) => { if (ends != null && day > ends) return; out.push({ day: day, rank: g.kind === 'inc' ? 0 : 1, idx: g.idx, kind: g.kind, desc: r.desc, inc: r.inc, bill: r.bill, note: r.note, src: r.r }); };
    if (step) {
      let d = days[days.length - 1] + step;
      while (d < start) d += step;
      for (; d <= end; d += step) add(d, lastRow);
    } else {
      g.rows.forEach(r => {
        const p = ymd_(r.day), lastDay = p.d === daysIn_(p.y, p.m);
        add(dayNum_(y, m, lastDay ? daysIn_(y, m) : Math.min(p.d, daysIn_(y, m))), r);
      });
    }
  });

  // weekly Everyday spending / Moved to savings rows, split on the same weekday as before (payday)
  const spendRows = last.rows.filter(r => kind_(r) === 'spend'), saveRows = last.rows.filter(r => kind_(r) === 'save');
  const wd = spendRows.length > 1 && spendRows[1].day != null ? ymd_(spendRows[1].day).w : 5;
  const lab = n => { const p = ymd_(n); return MON_[p.m] + ' ' + p.d; };
  const dfn = n => { const p = ymd_(n); return 'DATE(' + p.y + ',' + (p.m + 1) + ',' + p.d + ')'; };
  for (let s = start; s <= end;) {
    let nxt = s + 1; while (ymd_(nxt).w !== wd) nxt++;
    const e = Math.min(nxt - 1, end);
    if (spendRows.length) out.push({ day: s, rank: 2, idx: 0, kind: 'spend', desc: 'Everyday spending (' + lab(s) + '–' + lab(e) + ')', note: spendRows[0].note, src: spendRows[0].r,
      billF: "=SUMIFS('Spending Log'!$C$2:$C,'Spending Log'!$A$2:$A,\">=\"&" + dfn(s) + ",'Spending Log'!$A$2:$A,\"<\"&" + dfn(e + 1) + ')' });
    if (saveRows.length) out.push({ day: s, rank: 3, idx: 0, kind: 'save', desc: 'Moved to savings (' + lab(s) + '–' + lab(e) + ')', note: saveRows[0].note, src: saveRows[0].r,
      billF: "=SUMIFS('Savings Log'!$C$2:$C,'Savings Log'!$A$2:$A,\">=\"&" + dfn(s) + ",'Savings Log'!$A$2:$A,\"<\"&" + dfn(e + 1) + ')',
      hF: "=$B$5+SUMIFS('Savings Log'!$C$2:$C,'Savings Log'!$A$2:$A,\">=\"&" + dfn(start) + ",'Savings Log'!$A$2:$A,\"<\"&" + dfn(e + 1) + ')' });
    s = e + 1;
  }
  out.sort((a, b) => a.day - b.day || a.rank - b.rank || a.idx - b.idx);

  const startRow = last.rows[0];
  const rows = [{ day: start, kind: 'start', desc: startRow.desc || 'Starting balance', src: startRow.r }].concat(out);
  const first = 13, L = first + rows.length - 1;
  const grid = rows.map((x, i) => {
    const r = first + i, p = ymd_(x.day);
    return {
      day: x.day, src: x.src,
      cells: [null, DOW_[p.w], x.desc,
        x.kind === 'inc' ? x.inc : '',
        x.kind === 'bill' ? x.bill : (x.billF || ''),
        x.kind === 'start' ? '=B4' : '=F' + (r - 1) + '+D' + r + '-E' + r,
        x.kind === 'start' ? '' : (x.note || ''),
        x.hF || '', '']
    };
  });
  const title = MON_[m] + ' ' + y, mStart = 'DATE(' + y + ',' + (m + 1) + ',1)', mEnd = 'DATE(' + y + ',' + (m + 1) + ',' + daysIn_(y, m) + ')';
  return {
    title: title, first: first, L: L, grid: grid,
    a1: last.a1.indexOf(MONTH_NAME_[last.m] + ' ' + last.y) >= 0 ? last.a1.replace(MONTH_NAME_[last.m] + ' ' + last.y, MONTH_NAME_[m] + ' ' + y) : last.a1.replace(/ — .*/, '') + ' — ' + MONTH_NAME_[m] + ' ' + y,
    header: [
      ['Starting balance (from ' + last.title + ')', "='" + last.title + "'!E8", '', 'Total income', '=SUM(D' + first + ':D' + L + ')'],
      ['Savings starting balance (from ' + last.title + ')', "='" + last.title + "'!E10", '', 'Total bills', '=SUM(E' + first + ':E' + L + ')-E6-E9'],
      ['', '', '', 'Everyday spending', "=SUMIFS('Spending Log'!$C$2:$C,'Spending Log'!$A$2:$A,\">=\"&" + mStart + ",'Spending Log'!$A$2:$A,\"<=\"&" + mEnd + ')'],
      ['', '', '', 'Lowest balance', '=MIN(F' + first + ':F' + L + ')'],
      ['', '', '', 'Ending balance', '=F' + L],
      ['', '', '', 'Moved to savings (net)', "=SUMIFS('Savings Log'!$C$2:$C,'Savings Log'!$A$2:$A,\">=\"&" + mStart + ",'Savings Log'!$A$2:$A,\"<=\"&" + mEnd + ')'],
      ['', '', '', 'Savings ending balance', '=B5+E9']
    ],
    note: 'Notes: Starting balances carry over from ' + last.title + '. This month was added automatically from ' + last.title + ' — monthly bills on the same day, weekly ones on the same weekday. Edit anything that changed.'
  };
}

function makeNextMonth_(ss, lastT, prevT) {
  const tz = ss.getSpreadsheetTimeZone();
  const last = readMonth_(lastT, tz), prev = prevT ? readMonth_(prevT, tz) : null;
  const y = lastT.m === 11 ? lastT.y + 1 : lastT.y, m = (lastT.m + 1) % 12;
  const plan = planMonth_(last, prev, y, m);
  if (ss.getSheetByName(plan.title)) throw new Error(plan.title + ' already exists');

  const sh = lastT.sh.copyTo(ss).setName(plan.title);
  ss.setActiveSheet(sh); ss.moveActiveSheet(ss.getNumSheets());
  const need = plan.L + 4;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  sh.getRange(plan.first, 1, sh.getMaxRows() - plan.first + 1, 9).clear();

  sh.getRange(1, 1).setValue(plan.a1);
  sh.getRange(4, 1, plan.header.length, 5).setValues(plan.header);
  const toDate = n => { const p = ymd_(n); return Utilities.parseDate((p.m + 1) + '/' + p.d + '/' + p.y, tz, 'M/d/yyyy'); };
  sh.getRange(plan.first, 1, plan.grid.length, 9).setValues(plan.grid.map(g => { const c = g.cells.slice(); c[0] = toDate(g.day); return c; }));
  // same look as the matching rows in last month
  plan.grid.forEach((g, i) => lastT.sh.getRange(g.src, 1, 1, 9).copyTo(sh.getRange(plan.first + i, 1, 1, 9), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false));
  sh.getRange(plan.L + 2, 1).setValue(plan.note);
  SpreadsheetApp.flush();
  return { sh: sh, title: plan.title, y: y, m: m, key: y * 12 + m };
}

// Moves month tabs and log entries from before January of last year into "<name> Archive"
function archiveOld_(ss, name) {
  const cutY = new Date().getFullYear() - KEEP_YEARS_BACK;
  const tabs = monthTabs_(ss), old = tabs.filter(t => t.y < cutY), keep = tabs.filter(t => t.y >= cutY);
  if (!old.length || !keep.length) return [];
  const tz = ss.getSpreadsheetTimeZone(), arch = archiveFor_(ss, name);

  // 1) copy the old months into the archive with their numbers frozen as they were, then check every cell came across
  old.forEach(t => {
    const copy = t.sh.copyTo(arch);
    let nm = t.title, n = 2; while (arch.getSheetByName(nm)) nm = t.title + ' (' + n++ + ')';
    copy.setName(nm);
    const src = t.sh.getDataRange().getValues();
    copy.getRange(1, 1, src.length, src[0].length).setValues(src);
    SpreadsheetApp.flush();
    if (!same_(src, copy.getRange(1, 1, src.length, src[0].length).getValues())) throw new Error('Archive copy of ' + t.title + " didn't match, so nothing was removed.");
  });

  // 2) copy the old log entries and check them
  const cutoff = Utilities.parseDate('1/1/' + cutY, tz, 'M/d/yyyy');
  const logs = LOGS_.map(n => copyOldLog_(ss.getSheetByName(n), arch, n, cutoff)).filter(Boolean);

  // 3) everything is safely in the archive: the first month that stays keeps its starting balances as plain numbers
  const k = keep[0].sh;
  k.getRange(4, 2, 2, 1).setValues(k.getRange(4, 2, 2, 1).getValues());
  k.getRange(4, 1, 2, 1).setValues([['Starting balance (carried over from ' + old[old.length - 1].title + ')'], ['Savings starting balance (carried over from ' + old[old.length - 1].title + ')']]);

  // 4) now remove the old entries and months from the main sheet
  logs.forEach(x => { x.sh.getRange(2, 1, x.rows, x.width).clearContent(); if (x.kept.length) x.sh.getRange(2, 1, x.kept.length, x.width).setValues(x.kept); });
  old.forEach(t => ss.deleteSheet(t.sh));
  const blank = arch.getSheetByName('Sheet1'); if (blank && arch.getSheets().length > 1 && blank.getLastRow() === 0) arch.deleteSheet(blank);
  return old.map(t => t.title);
}

// cells match (dates compared by time, numbers to the cent)
function same_(a, b) {
  if (a.length !== b.length) return false;
  for (let r = 0; r < a.length; r++) for (let c = 0; c < a[r].length; c++) {
    const x = a[r][c], y = b[r] ? b[r][c] : undefined;
    if (x instanceof Date || y instanceof Date) { if (!(x instanceof Date && y instanceof Date) || x.getTime() !== y.getTime()) return false; }
    else if (typeof x === 'number' || typeof y === 'number') { if (Math.abs(Number(x) - Number(y)) > 0.004) return false; }
    else if (String(x) !== String(y)) return false;
  }
  return true;
}

// Copies log entries from before the cutoff into the archive and checks them. Removing them from the main sheet happens later.
function copyOldLog_(sh, arch, name, cutoff) {
  if (!sh || sh.getLastRow() < 2) return null;
  const width = sh.getLastColumn(), data = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
  const old = data.filter(r => r[0] instanceof Date && r[0] < cutoff), kept = data.filter(r => !(r[0] instanceof Date && r[0] < cutoff));
  if (!old.length) return null;
  let a = arch.getSheetByName(name);
  if (!a) { a = arch.insertSheet(name); a.getRange(1, 1, 1, width).setValues(sh.getRange(1, 1, 1, width).getValues()); a.setFrozenRows(1); }
  const at = a.getLastRow() + 1;
  a.getRange(at, 1, old.length, width).setValues(old);
  SpreadsheetApp.flush();
  if (!same_(old, a.getRange(at, 1, old.length, width).getValues())) throw new Error('Archive copy of the ' + name + " didn't match, so nothing was removed.");
  return { sh: sh, rows: data.length, width: width, kept: kept };
}

function archiveFor_(ss, name) {
  const props = PropertiesService.getScriptProperties(), key = 'archive:' + ss.getId(), id = props.getProperty(key);
  if (id) { try { return SpreadsheetApp.openById(id); } catch (err) { /* archive was removed: make a new one */ } }
  const arch = SpreadsheetApp.create(name + ' Archive');
  props.setProperty(key, arch.getId());
  return arch;
}
