// Garage – Google Apps Script backend, version 8.
// Sheets sync, Calendar reminders, Maps location, Drive receipts, daily Drive backups, an error log,
// and (new in 7) a read-only Finance link: the car loan balance and insurance sinking fund from your
// Finance (Ledger) Google Sheet, and (new in 8) receipt scanning with Google Drive's text recognition.
//
// Setup after pasting this file:
// 1) Put your own long secret word in TOKEN below (the same one you type in Garage > Settings).
//    Tip: Garage > Settings > "Make a strong secret word" makes one for you.
// 2) For the Finance link: Project Settings > Script Properties > Add property
//    FINANCE_SHEET_ID = the ID of your Finance Sheet (the long part of its address between /d/ and /edit).
// 3) For receipt scanning: in the editor's left bar, Services (+) > Drive API > Add.
// 4) Choose "authorize" in the toolbar and click Run once (it asks to allow the new access).
// 5) Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
const TOKEN = 'CHANGE_ME';
const VERSION = 8;
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

function authorize() {
  LockService.getScriptLock(); CacheService.getScriptCache(); SpreadsheetApp.getActiveSpreadsheet();
  CalendarApp.getDefaultCalendar(); Maps.newGeocoder().reverseGeocode(3.139, 101.687); DriveApp.getRootFolder();
  const fin = PropertiesService.getScriptProperties().getProperty('FINANCE_SHEET_ID');
  if (fin) Logger.log('Finance Sheet: ' + SpreadsheetApp.openById(fin).getName());
  else Logger.log('FINANCE_SHEET_ID is not set yet (only needed for the Finance link).');
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out({ error: 'Bad request' }); }
  const cache = CacheService.getScriptCache();
  const fails = +(cache.get('fails') || 0);
  if (fails >= 20) return out({ error: 'Too many wrong secret words. Locked for 10 minutes.' });
  if (TOKEN === 'CHANGE_ME') return out({ error: 'Set your secret word in the first lines of Code.gs' });
  if (typeof b.token !== 'string' || b.token !== TOKEN) {
    cache.put('fails', String(fails + 1), 600);
    logErr('auth', 'Wrong secret word');
    return out({ error: 'Wrong token' });
  }
  try { return out(route(b)); }
  catch (x) { logErr(b.action || 'sync', (x && x.stack) || x); return out({ error: String((x && x.message) || x) }); }
}

function route(b) {
  switch (b.action) {
    case 'geo': return geo(b);
    case 'up': return upload(b);
    case 'get': return getReceipt(b.fid);
    case 'check': return checkReceipts(b.rids);
    case 'health': return health();
    case 'backups': return listBackups();
    case 'getbackup': return getBackup(b.fid);
    case 'finance': return garageFinance_();
    case 'ocr': return ocr(b);
    case 'snap': return withLock(() => ({ name: saveBackup(Object.values(readData(dataSheet())), 'before-restore') }));
    case undefined: case null: case '': case 'sync': return withLock(() => syncAll(b));
    default: return { error: 'Unknown action: ' + String(b.action).slice(0, 40) };
  }
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return { error: 'Busy, try again' };
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---------- Sheets sync ----------
function dataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().find(s => s.getName() !== 'Log') || ss.insertSheet('Data', 0);
}
function readData(sh) {
  const m = {};
  sh.getDataRange().getValues().forEach(r => {
    if (!r[0] || r[0] === 'id') return;
    try { m[r[0]] = JSON.parse(r[1]); } catch (x) { logErr('sheet', 'Skipped an unreadable row: ' + r[0]); }
  });
  return m;
}
function syncAll(b) {
  const sh = dataSheet(), m = readData(sh);
  (Array.isArray(b.items) ? b.items : []).forEach(i => {
    if (!i || typeof i.id !== 'string' || !/^[\w.:-]{1,80}$/.test(i.id)) return;
    if (!m[i.id] || (+m[i.id].updated || 0) < (+i.updated || 0)) m[i.id] = i;
  });
  const all = Object.values(m);
  const trash = id => { try { const f = DriveApp.getFileById(id); if (inReceipts(f)) f.setTrashed(true); } catch (x) {} };
  all.forEach(i => { if (i.deleted && i.rid) { trash(i.rid); i.rid = ''; } });
  (Array.isArray(b.drop) ? b.drop : []).forEach(trash);
  sh.clear();
  sh.appendRow(['id', 'json (do not edit)', 'type', 'date', 'odometer', 'item', 'cost', 'notes', 'deleted']);
  if (all.length) sh.getRange(2, 1, all.length, 9).setValues(all.map(i =>
    [i.id, JSON.stringify(i), i.kind, i.date || '', i.odo || '', i.item || i.name || '', i.cost || '', i.notes || '', i.deleted ? 'yes' : '']));
  let cal = '';
  if (b.events) { try { syncCal(b.events); cal = 'ok'; } catch (x) { cal = String((x && x.message) || x); logErr('calendar', x); } }
  let bk = '';
  try { bk = autoBackup(all); } catch (x) { logErr('backup', x); }
  return { items: all, cal: cal, bk: bk, v: VERSION };
}

// ---------- Calendar ----------
function syncCal(ev) {
  const cal = CalendarApp.getCalendarsByName('Car Log')[0] || CalendarApp.createCalendar('Car Log');
  const P = PropertiesService.getScriptProperties();
  const map = JSON.parse(P.getProperty('cal') || '{}'), seen = {};
  (Array.isArray(ev) ? ev : []).forEach(e => {
    if (!e || !/^\d{4}-\d{2}-\d{2}$/.test(e.date)) return;
    const title = String(e.title || '').slice(0, 200), key = String(e.key || '').slice(0, 120);
    seen[key] = 1;
    const m = map[key];
    if (m && m.d === e.date && m.t === title) return;
    if (m) { try { cal.getEventById(m.id).deleteEvent(); } catch (x) {} }
    const x = cal.createAllDayEvent(title, new Date(e.date + 'T00:00:00'), { description: 'Added by Car Log' });
    x.removeAllReminders(); x.addPopupReminder(7 * 24 * 60); x.addPopupReminder(24 * 60);
    map[key] = { id: x.getId(), d: e.date, t: title };
  });
  Object.keys(map).forEach(k => {
    if (!seen[k]) { try { cal.getEventById(map[k].id).deleteEvent(); } catch (x) {} delete map[k]; }
  });
  P.setProperty('cal', JSON.stringify(map));
}

// ---------- Maps ----------
function geo(b) {
  const lat = +b.lat, lng = +b.lng;
  if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) return { error: 'Bad location' };
  const r = Maps.newGeocoder().setLanguage('en').setRegion('my').reverseGeocode(lat, lng);
  const top = r.results[0] || {}, c = top.address_components || [];
  const f = t => { const x = c.find(a => a.types.indexOf(t) > -1); return x ? x.long_name : ''; };
  return { geo: 1, state: f('administrative_area_level_1'), area: f('locality') || f('sublocality_level_1') || f('administrative_area_level_2'), address: top.formatted_address || '' };
}

// ---------- Drive receipts ----------
function rootFolder() { return topFolder('Garage Receipts'); }
function topFolder(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
function sub(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
// Garage Receipts / 2026 / 10 October / file
function folder(name) {
  const root = rootFolder();
  const m = /^(\d{4})-(\d{2})/.exec(name || '');
  const mi = m ? parseInt(m[2], 10) : 0;
  if (!m || mi < 1 || mi > 12) return root;
  return sub(sub(root, m[1]), m[2] + ' ' + MONTHS[mi - 1]);
}
function upload(b) {
  const mime = String(b.mime || '');
  if (!/^(image\/(jpeg|png|webp|heic|heif)|application\/pdf)$/.test(mime)) return { error: 'File type not allowed' };
  const data = String(b.data || '');
  if (data.length > 12e6) return { error: 'File too large' };
  const name = String(b.name || 'receipt').replace(/[^\w ,.()-]/g, '').slice(0, 80) || 'receipt';
  const f = folder(name).createFile(Utilities.newBlob(Utilities.base64Decode(data), mime, name));
  return { fid: f.getId(), url: f.getUrl() };
}
function getReceipt(fid) {
  const f = DriveApp.getFileById(String(fid));
  if (!inReceipts(f)) return { error: 'Not a Garage receipt' };
  if (f.isTrashed()) f.setTrashed(false);
  return { mime: f.getMimeType(), data: Utilities.base64Encode(f.getBlob().getBytes()) };
}
function checkReceipts(rids) {
  const missing = [];
  (Array.isArray(rids) ? rids : []).slice(0, 500).forEach(id => {
    try { const f = DriveApp.getFileById(String(id)); if (f.isTrashed() || !inReceipts(f)) missing.push(id); } catch (x) { missing.push(id); }
  });
  return { missing: missing };
}
// Only touch files that live inside Garage Receipts (or its year/month folders).
function inReceipts(f) {
  const rootId = rootFolder().getId();
  let ps = f.getParents();
  for (let depth = 0; depth < 4 && ps.hasNext(); depth++) {
    const p = ps.next();
    if (p.getId() === rootId) return true;
    ps = p.getParents();
  }
  return false;
}
// Run once from the editor to tidy files that were uploaded before year/month folders existed.
function organize() {
  const root = rootFolder(), files = root.getFiles(), list = [];
  while (files.hasNext()) list.push(files.next());
  let n = 0;
  list.forEach(f => {
    const t = folder(f.getName());
    if (t.getId() !== root.getId()) { f.moveTo(t); n++; }
  });
  Logger.log('Moved ' + n + ' files');
}

// ---------- Daily backups: Drive / Garage Backups ----------
function backupFolder() { return topFolder('Garage Backups'); }
function stamp(fmt) { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), fmt); }
// Refreshes today's backup at most every 30 minutes.
function autoBackup(all) {
  const P = PropertiesService.getScriptProperties();
  if (Date.now() - +(P.getProperty('bkAt') || 0) < 30 * 60 * 1000) return P.getProperty('bkDay') || '';
  saveBackup(all, '');
  const day = stamp('yyyy-MM-dd');
  P.setProperty('bkAt', String(Date.now())); P.setProperty('bkDay', day);
  return day;
}
function saveBackup(all, label) {
  const fo = backupFolder();
  const name = 'garage-backup-' + stamp('yyyy-MM-dd') + (label ? '-' + label + '-' + stamp('HHmm') : '') + '.json';
  const body = JSON.stringify({ v: 2, app: 'Garage', saved: new Date().toISOString(), items: all });
  const it = fo.getFilesByName(name);
  if (it.hasNext()) it.next().setContent(body); else fo.createFile(name, body, 'application/json');
  prune(fo);
  return name;
}
// Keeps the 40 newest backups; older ones go to the Drive bin.
function prune(fo) {
  const files = [], it = fo.getFiles();
  while (it.hasNext()) { const f = it.next(); if (/^garage-backup-.*\.json$/.test(f.getName())) files.push(f); }
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated()).slice(40).forEach(f => f.setTrashed(true));
}
function listBackups() {
  const fo = backupFolder(), it = fo.getFiles(), o = [];
  while (it.hasNext()) {
    const f = it.next();
    if (/^garage-backup-.*\.json$/.test(f.getName())) o.push({ fid: f.getId(), name: f.getName(), at: f.getLastUpdated().getTime(), size: f.getSize() });
  }
  o.sort((a, b) => b.at - a.at);
  return { backups: o.slice(0, 40) };
}
function getBackup(fid) {
  const f = DriveApp.getFileById(String(fid)), ps = f.getParents();
  if (!ps.hasNext() || ps.next().getId() !== backupFolder().getId()) return { error: 'Not a Garage backup' };
  return { data: f.getBlob().getDataAsString() };
}

// ---------- Receipt scanning ----------
// Turns the photo into a temporary Google Doc with Drive's text recognition, reads the text, then bins the Doc.
// Needs the Drive API advanced service (Services > + > Drive API). The app picks out the amounts itself.
function ocr(b) {
  const mime = String(b.mime || '');
  if (!/^(image\/(jpeg|png)|application\/pdf)$/.test(mime)) return { error: 'File type not allowed' };
  const data = String(b.data || '');
  if (!data || data.length > 8e6) return { error: 'File too large' };
  if (typeof Drive === 'undefined') return { text: '', ocrError: 'Turn on Drive API in Apps Script (Services > + > Drive API), then deploy a new version' };
  const blob = Utilities.newBlob(Utilities.base64Decode(data), mime, 'garage-scan');
  let id;
  if (Drive.Files.create) id = Drive.Files.create({ name: 'garage-scan', mimeType: 'application/vnd.google-apps.document' }, blob, { ocrLanguage: 'en', fields: 'id' }).id;
  else id = Drive.Files.insert({ title: 'garage-scan', mimeType: 'application/vnd.google-apps.document' }, blob, { ocr: true, ocrLanguage: 'en' }).id;
  try { return { text: DocumentApp.openById(id).getBody().getText().slice(0, 20000) }; }
  finally { try { DriveApp.getFileById(id).setTrashed(true); } catch (x) {} }
}

// ---------- Finance link (read-only) ----------
// Reads the readable "Accounts" and "Goals" tabs that the Finance (Ledger) app writes, never its
// private _data tab, and never writes anything. Mirrors the link the other way (Ledger reads this
// Sheet via GARAGE_ID). Garage shows the figures but does not store them.
function garageFinance_() {
  const id = PropertiesService.getScriptProperties().getProperty('FINANCE_SHEET_ID');
  if (!id) return { fin: null, finError: 'Add FINANCE_SHEET_ID to Garage\'s Script Properties' };
  let ss;
  try { ss = SpreadsheetApp.openById(id); } catch (x) { return { fin: null, finError: 'Cannot open the Finance Sheet: ' + ((x && x.message) || x) }; }
  const acc = finTable_(ss, 'Accounts'), goals = finTable_(ss, 'Goals');
  return {
    fin: {
      at: new Date().toISOString(),
      accounts: acc.map(r => {
        const owed = finNum_(r['owed']);
        return { name: finText_(r['account']), type: finText_(r['type']).toLowerCase(), balance: finNum_(r['balance']), owed: owed === null ? null : Math.abs(owed) };
      }).filter(a => a.name),
      goals: goals.map(r => ({
        name: finText_(r['goal']), type: finText_(r['type']), target: finNum_(r['target']), saved: finNum_(r['set aside (manual goals)']),
        due: finDate_(r['target date']), achieved: /^(true|yes|1)$/i.test(finText_(r['achieved']))
      })).filter(g => g.name)
    }
  };
}
// Rows of a tab as objects keyed by lower-case header (row 1). Missing tab -> [].
function finTable_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const head = v[0].map(h => String(h).trim().toLowerCase());
  return v.slice(1).map(row => { const o = {}; head.forEach((h, i) => { if (h) o[h] = row[i]; }); return o; });
}
function finText_(x) { return x === null || x === undefined ? '' : String(x).trim(); }
function finNum_(x) {
  if (x === '' || x === null || x === undefined) return null;
  const n = typeof x === 'number' ? x : Number(String(x).replace(/[^0-9.\-]/g, ''));
  return isFinite(n) ? Math.round(n * 100) / 100 : null;
}
function finDate_(x) {
  if (Object.prototype.toString.call(x) === '[object Date]') return Utilities.formatDate(x, 'Asia/Kuala_Lumpur', 'yyyy-MM-dd');
  const s = finText_(x);
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : '';
}

// ---------- Connection check ----------
function health() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sh = dataSheet(), P = PropertiesService.getScriptProperties();
  const rf = DriveApp.getFoldersByName('Garage Receipts'), bf = DriveApp.getFoldersByName('Garage Backups');
  const lg = ss.getSheetByName('Log'), n = lg ? Math.max(0, lg.getLastRow() - 1) : 0;
  let finance = null;
  const fid = P.getProperty('FINANCE_SHEET_ID');
  if (fid) { try { const f = SpreadsheetApp.openById(fid); finance = { name: f.getName(), url: f.getUrl() }; } catch (x) { finance = { error: String((x && x.message) || x) }; } }
  return {
    v: VERSION,
    sheet: { name: ss.getName(), url: ss.getUrl(), rows: Math.max(0, sh.getLastRow() - 1) },
    receipts: rf.hasNext() ? { url: rf.next().getUrl() } : null,
    backups: bf.hasNext() ? { url: bf.next().getUrl(), last: P.getProperty('bkDay') || '' } : null,
    calendar: !!CalendarApp.getCalendarsByName('Car Log')[0],
    finance: finance,
    ocr: typeof Drive !== 'undefined',
    errors: n,
    recent: n ? lg.getRange(Math.max(2, lg.getLastRow() - 4), 1, Math.min(5, n), 3).getValues().map(r => [new Date(r[0]).getTime(), String(r[1]), String(r[2]).slice(0, 200)]).reverse() : []
  };
}

// ---------- Error log: a "Log" tab in this Google Sheet ----------
function logErr(area, msg) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('Log');
    if (!sh) { sh = ss.insertSheet('Log', ss.getNumSheets()); sh.appendRow(['time', 'area', 'message']); }
    sh.appendRow([new Date(), String(area), String(msg).slice(0, 500)]);
    if (sh.getLastRow() > 600) sh.deleteRows(2, sh.getLastRow() - 500);
  } catch (x) {}
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
