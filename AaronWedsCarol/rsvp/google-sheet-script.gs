/**
 * Records RSVPs from the Aaron & Carol website in this Google Sheet.
 *
 * Setup (once):
 *  1. Create a Google Sheet, e.g. "Aaron & Carol RSVPs".
 *  2. In the sheet: Extensions > Apps Script. Delete what is there and paste
 *     this whole file. Save.
 *  3. Deploy > New deployment > Select type: Web app.
 *       Execute as: Me
 *       Who has access: Anyone
 *     Click Deploy, allow the permissions it asks for, and copy the Web app URL
 *     (it ends in /exec).
 *  4. Put that URL in index.html:  const RSVP_SHEET="https://script.google.com/macros/s/…/exec";
 *
 * If you edit this script later, use Deploy > Manage deployments > Edit
 * (pencil) > Version: New version, so the same URL keeps working.
 */

const HEADERS = ['Received', 'Name', 'Reply', 'Guests', 'Note', 'Invited as', 'Sent from page'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);              // two guests replying at once never overwrite each other
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    const r = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    sheet.appendRow([
      new Date(),
      clean(r.name),
      clean(r.attending),
      clean(r.guests),
      clean(r.note),
      clean(r.invitedAs),
      clean(r.sentAt),
    ]);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

// Keep text from the form as plain text: a leading = + - or @ would otherwise
// be read by Sheets as a formula.
function clean(v) {
  const s = String(v == null ? '' : v).slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// Opening the URL in a browser shows this, which is a quick way to check it is live.
function doGet() {
  return ContentService.createTextOutput('Aaron & Carol RSVP sheet is ready.');
}
