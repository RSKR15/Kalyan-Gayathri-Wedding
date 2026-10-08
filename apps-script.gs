/**
 * Wedding RSVP — Google Apps Script
 * Paste this into Extensions → Apps Script of your RSVP Google Sheet,
 * then Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 *
 * One row per phone number: if a guest RSVPs again with the same number,
 * their row is updated instead of duplicated.
 */
const SHEET_NAME = 'RSVPs';
const HEADERS = ['Last updated', 'Name', 'Phone', 'Attending', 'Mehendi', 'Haldi & Ceremonies', 'Wedding', 'Guests', 'Note'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('C:C').setNumberFormat('@'); // keep phone numbers as text
  }
  migrateColumns_(sh);
  // keep the header row up to date (new or renamed columns)
  const head = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (head.join('|') !== HEADERS.join('|')) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  }
  return sh;
}

// One-time: move the old column order (… Mehendi, Wedding, Guests, Note, Nov 19)
// to the new order (… Mehendi, Haldi & Ceremonies, Wedding, Guests, Note) and use "-" for events not on that invite.
function migrateColumns_(sh) {
  if (sh.getLastColumn() < 6 || sh.getRange(1, 6).getValue() !== 'Wedding') return;
  const dash = v => (v === '' || v === 'wedding link') ? '-' : v;
  const last = sh.getLastRow();
  if (last >= 2) {
    const rows = sh.getRange(2, 1, last - 1, 9).getValues().map(r =>
      [r[0], r[1], r[2], r[3], dash(r[4]), dash(r[8]), dash(r[5]), r[6], r[7]]);
    sh.getRange(2, 1, rows.length, 9).setValues(rows);
  }
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
}

const digits_ = s => String(s || '').replace(/\D/g, '');

function findRow_(sh, phone) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const phones = sh.getRange(2, 3, last - 1, 1).getValues();
  for (let i = 0; i < phones.length; i++) {
    if (digits_(phones[i][0]) === phone) return i + 2;
  }
  return -1;
}

function doPost(e) {
  const p = e.parameter;
  // photos don't need to wait behind RSVP saves — handle them first, without the lock
  if (p.action === 'memory') return saveMemory_(p);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const phone = digits_(p.phone);
    if (!phone) return json_({ result: 'error', error: 'missing phone' });
    const sh = getSheet_();
    const row = [new Date(), p.name || '', phone, p.attending || '', p.mehendi || '-',
                 p.day2 || '-', p.wedding || '-', Number(p.guests || 0), p.note || ''];
    const r = findRow_(sh, phone);
    if (r > 0) sh.getRange(r, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return json_({ result: 'success' });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const phone = digits_(e.parameter.phone);
  if (!phone) return json_({ result: 'error', error: 'missing phone' });
  const sh = getSheet_();
  const r = findRow_(sh, phone);
  if (r < 0) return json_({ result: 'not_found' });
  const v = sh.getRange(r, 1, 1, HEADERS.length).getValues()[0];
  return json_({
    result: 'success',
    data: { name: v[1], phone: String(v[2]), attending: v[3], mehendi: v[4],
            day2: v[5], wedding: v[6], guests: String(v[7]), note: v[8] }
  });
}

/* ---------- Share a Memory: photos saved to Drive + logged in a "Memories" tab ---------- */
const MEM_FOLDER = 'Wedding Memories – Sai Kalyan & Gayathri';
const MEM_SHEET = 'Memories';

function memFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('MEM_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (err) {} }
  const it = DriveApp.getFoldersByName(MEM_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(MEM_FOLDER);
  props.setProperty('MEM_FOLDER_ID', folder.getId());
  return folder;
}

function memSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MEM_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MEM_SHEET);
    sh.appendRow(['Shared on', 'Name', 'Their words', 'Photo']);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Run this once from the editor (Run ▶ setupMemories) to allow Drive access and create the folder + Memories tab.
function setupMemories() {
  const folder = memFolder_();
  memSheet_();
  Logger.log('Memories folder ready: ' + folder.getUrl());
}

function saveMemory_(p) {
  const mime = String(p.mime || 'image/jpeg');
  if (!/^image\//.test(mime) || !p.data) return json_({ result: 'error', error: 'images only' });
  const bytes = Utilities.base64Decode(p.data);
  if (bytes.length > 8 * 1024 * 1024) return json_({ result: 'error', error: 'too large' });
  const name = String(p.name || 'Guest').replace(/[\\/:*?"<>|]/g, '').trim().slice(0, 60) || 'Guest';
  const note = String(p.note || '').slice(0, 500);
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmmss');
  const fileName = name + ' – ' + stamp + ' – ' + (p.idx || 1) + '.jpg';
  const file = memFolder_().createFile(Utilities.newBlob(bytes, mime, fileName));
  if (note) file.setDescription(note);
  memSheet_().appendRow([new Date(), name, note, file.getUrl()]);
  return json_({ result: 'success' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
