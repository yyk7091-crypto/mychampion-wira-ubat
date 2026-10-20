/**
 * Backend for "Cabaran Wira Ubat MyCHAMPION 2026".
 *
 * SETUP (once):
 *   1. Create a Google Sheet. Extensions > Apps Script. Paste this whole file into Code.gs.
 *   2. Run the `setup` function once (pick it in the toolbar, press Run) and approve access.
 *      It creates the "Peserta" tab with headers.
 *   3. Deploy > New deployment > type "Web app".
 *        Execute as:      Me
 *        Who has access:  Anyone        (NOT "Anyone with Google account": the game can't sign in)
 *   4. Copy the Web app URL (ends in /exec) into GOOGLE_APPS_SCRIPT_URL in the HTML file.
 *   5. Test: open the /exec URL in a browser. It should say "Wira Ubat backend OK".
 *
 * After editing this script you must Deploy > Manage deployments > Edit > Version: New version,
 * otherwise the old code keeps running at the same URL.
 *
 * AT THE COUNTER: press Ctrl+F in the Sheet, search the voucher code (or phone number),
 * then tick the "Ditebus" checkbox. The redemption time fills in automatically.
 */

const SHEET_NAME = 'Peserta';
const HEADERS = ['Masa Daftar', 'Kod Baucar', 'Nama', 'No. Telefon / E-mel', 'Skor', 'Skor Maksimum',
                 'Persetujuan', 'Sumber', 'Ditebus', 'Masa Ditebus'];
const COL = { code: 2, contact: 4, redeemed: 9, redeemedAt: 10 }; // 1-based column numbers
const CODE_CHARS = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
const CODE_PATTERN = /^MC-[A-HJ-NP-Z2-9]{5}$/;

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    // Only one registration is written at a time, so two players never get the same code.
    lock.waitLock(20000);

    const data = JSON.parse(e.postData.contents);
    const name = clean_(data.name);
    const contact = clean_(data.contact);
    if (!name || !contact || data.consent !== true) {
      return json_({ status: 'error', message: 'Maklumat tidak lengkap' });
    }

    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues() : [];

    // One voucher per phone number / e-mail: return the original code instead of a new one.
    const key = normaliseContact_(contact);
    const existing = rows.find(function (r) { return normaliseContact_(r[COL.contact - 1]) === key; });
    if (existing) {
      return json_({
        status: 'success',
        serial: existing[COL.code - 1],
        existing: true,
        redeemed: existing[COL.redeemed - 1] === true
      });
    }

    // Keep the phone's offline code if it's valid and unused, so the code on screen matches the Sheet.
    const usedCodes = new Set(rows.map(function (r) { return r[COL.code - 1]; }));
    let serial = CODE_PATTERN.test(data.serial) && !usedCodes.has(data.serial) ? data.serial : null;
    while (!serial || usedCodes.has(serial)) serial = generateCode_();

    const maxScore = Number(data.maxScore) || 0;
    const score = Math.max(0, Math.min(Number(data.score) || 0, maxScore));

    sheet.appendRow([
      new Date(),
      serial,
      asText_(name),
      asText_(contact),
      score,
      maxScore,
      'Ya',
      data.offline ? 'Luar talian (dihantar kemudian)' : 'Dalam talian',
      false,
      ''
    ]);
    sheet.getRange(sheet.getLastRow(), COL.redeemed).insertCheckboxes();

    return json_({ status: 'success', serial: serial, existing: false });
  } catch (err) {
    console.error(err);
    return json_({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Lets you check the deployment by opening the /exec URL in a browser.
function doGet() {
  return ContentService.createTextOutput('Wira Ubat backend OK');
}

// Simple trigger: stamps the time when staff tick "Ditebus", clears it if they untick.
function onEdit(e) {
  const range = e.range;
  if (range.getSheet().getName() !== SHEET_NAME) return;
  if (range.getColumn() !== COL.redeemed || range.getRow() < 2 || range.getNumRows() > 1) return;
  range.offset(0, COL.redeemedAt - COL.redeemed).setValue(range.getValue() === true ? new Date() : '');
}

// Run once from the editor to create the tab and approve permissions.
function setup() {
  getSheet_();
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function generateCode_() {
  let code = 'MC-';
  for (let i = 0; i < 5; i++) code += CODE_CHARS.charAt(Math.floor(Math.random() * CODE_CHARS.length));
  return code;
}

function clean_(value) {
  return String(value == null ? '' : value).replace(/\s+/g, ' ').trim().slice(0, 100);
}

// Leading apostrophe stores the value as plain text, so phone numbers keep their leading 0,
// "+6012..." isn't read as a number, and text starting with "=" can't run as a formula.
function asText_(value) {
  return "'" + value;
}

function normaliseContact_(value) {
  const v = String(value).trim().toLowerCase();
  if (v.indexOf('@') !== -1) return v;
  let digits = v.replace(/\D/g, '');
  if (digits.indexOf('60') === 0) digits = digits.slice(1); // +6012... and 012... are the same number
  return digits;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
