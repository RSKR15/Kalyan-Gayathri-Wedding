/**
 * Wedding RSVP — Google Apps Script
 * Paste this into Extensions → Apps Script of your RSVP Google Sheet,
 * then Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 *
 * One row per phone number: if a guest RSVPs again with the same number,
 * their row is updated instead of duplicated.
 */
const SHEET_NAME = 'RSVPs';
const HEADERS = ['Last updated', 'Name', 'Phone', 'Attending', 'Mehendi', 'Wedding', 'Guests', 'Note', 'Nov 19 (Haldi & ceremonies)'];

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
  // add any new header columns to an existing sheet
  if (sh.getLastColumn() < HEADERS.length) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  }
  return sh;
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
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = e.parameter;
    const phone = digits_(p.phone);
    if (!phone) return json_({ result: 'error', error: 'missing phone' });
    const sh = getSheet_();
    const row = [new Date(), p.name || '', phone, p.attending || '', p.mehendi || '',
                 p.wedding || '', Number(p.guests || 0), p.note || '', p.day2 || ''];
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
            wedding: v[5], guests: String(v[6]), note: v[7], day2: v[8] || '' }
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
