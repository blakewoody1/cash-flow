/**
 * BACKUP COPY. The live copy is in the Cash Flow sheet → Extensions → Apps Script.
 * The KEY below is a placeholder on purpose — never commit your real key to this public repo.
 *
 * Cash Flow — Apps Script for your "Cash Flow" Google Sheet
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
    // several reads in one trip (the app uses this to open faster on a slow signal)
    if (req.tool === 'batch') return { result: { results: ((req.input || {}).calls || []).map(c => run_(c.tool, c.input || {}, true)) } };
    return run_(req.tool, req.input || {}, false);
  } finally {
    if (lock) lock.releaseLock();
  }
}
function run_(tool, i, readOnly) {
  try {
    const ss = sheetFor_(i.spreadsheetId);
    switch (tool) {
      case 'get_spreadsheet': return { result: getSpreadsheet_(ss, i) };
      case 'get_values': return { result: getValues_(ss, i) };
      case 'update_values': if (!readOnly) return { result: updateValues_(ss, i) }; break;
      case 'insert_dimension': if (!readOnly) return { result: insertDimension_(ss, i) }; break;
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
      properties: { title: ss.getName(), spreadsheetId: ss.getId(), features: ['batch'] },
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
    const add = (day, r) => out.push({ day: day, rank: g.kind === 'inc' ? 0 : 1, idx: g.idx, kind: g.kind, desc: r.desc, inc: r.inc, bill: r.bill, note: r.note, src: r.r });
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

  // 1) the first month that stays keeps its starting balances as plain numbers
  const k = keep[0].sh;
  k.getRange(4, 2, 2, 1).setValues(k.getRange(4, 2, 2, 1).getValues());
  k.getRange(4, 1, 2, 1).setValues([['Starting balance (carried over from ' + old[old.length - 1].title + ')'], ['Savings starting balance (carried over from ' + old[old.length - 1].title + ')']]);

  // 2) copy the old months into the archive with their numbers frozen as they were
  old.forEach(t => {
    const copy = t.sh.copyTo(arch);
    let nm = t.title, n = 2; while (arch.getSheetByName(nm)) nm = t.title + ' (' + n++ + ')';
    copy.setName(nm);
    const rg = t.sh.getDataRange();
    copy.getRange(1, 1, rg.getNumRows(), rg.getNumColumns()).setValues(rg.getValues());
  });

  // 3) move the old log entries
  const cutoff = Utilities.parseDate('1/1/' + cutY, tz, 'M/d/yyyy');
  ['Spending Log', 'Savings Log'].forEach(n => moveOldLog_(ss.getSheetByName(n), arch, n, cutoff));

  // 4) remove the old months from the main sheet
  old.forEach(t => ss.deleteSheet(t.sh));
  const blank = arch.getSheetByName('Sheet1'); if (blank && arch.getSheets().length > 1 && blank.getLastRow() === 0) arch.deleteSheet(blank);
  return old.map(t => t.title);
}

function moveOldLog_(sh, arch, name, cutoff) {
  if (!sh || sh.getLastRow() < 2) return;
  const width = sh.getLastColumn(), data = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
  const old = data.filter(r => r[0] instanceof Date && r[0] < cutoff), kept = data.filter(r => !(r[0] instanceof Date && r[0] < cutoff));
  if (!old.length) return;
  let a = arch.getSheetByName(name);
  if (!a) { a = arch.insertSheet(name); a.getRange(1, 1, 1, width).setValues(sh.getRange(1, 1, 1, width).getValues()); a.setFrozenRows(1); }
  a.getRange(a.getLastRow() + 1, 1, old.length, width).setValues(old);
  sh.getRange(2, 1, data.length, width).clearContent();
  if (kept.length) sh.getRange(2, 1, kept.length, width).setValues(kept);
}

function archiveFor_(ss, name) {
  const props = PropertiesService.getScriptProperties(), key = 'archive:' + ss.getId(), id = props.getProperty(key);
  if (id) { try { return SpreadsheetApp.openById(id); } catch (err) { /* archive was removed: make a new one */ } }
  const arch = SpreadsheetApp.create(name + ' Archive');
  props.setProperty(key, arch.getId());
  return arch;
}
