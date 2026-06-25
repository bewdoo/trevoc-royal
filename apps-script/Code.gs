/**
 * TREVOC Royal Residences — Lead capture endpoint
 * Saves every website enquiry to a Google Sheet AND emails a notification.
 *
 * SETUP (one time):
 *  1. Create a new Google Sheet (e.g. "TREVOC Leads").
 *  2. In that Sheet: Extensions ▸ Apps Script.
 *  3. Delete the default code, paste THIS file, and Save.
 *  4. Deploy ▸ New deployment ▸ type "Web app".
 *       - Execute as:  Me
 *       - Who has access:  Anyone
 *     Deploy, authorise, and COPY the "/exec" Web app URL.
 *  5. Paste that URL into LEAD_ENDPOINT in index.html (see the website code).
 *
 * To change where emails go, edit NOTIFY_EMAIL below.
 */

var NOTIFY_EMAIL = 'unplugged.realty11@gmail.com';
var SHEET_NAME   = 'Leads';

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    var p = (e && e.parameter) ? e.parameter : {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Name', 'Phone', 'Email', 'Residence', 'Message', 'Source', 'Page']);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var ts = new Date();
    sheet.appendRow([
      ts,
      p.name    || '',
      p.phone   || '',
      p.email   || '',
      p.config  || '',
      p.message || '',
      p.source  || '',
      p.page    || ''
    ]);

    var subject = 'New TREVOC Lead — ' + (p.name || 'Enquiry') + ' · ' + (p.config || 'General');
    var body =
      'A new enquiry just came in from the TREVOC Royal Residences website.\n\n' +
      'Name:      ' + (p.name    || '—') + '\n' +
      'Phone:     ' + (p.phone   || '—') + '\n' +
      'Email:     ' + (p.email   || '—') + '\n' +
      'Residence: ' + (p.config  || '—') + '\n' +
      'Message:   ' + (p.message || '—') + '\n' +
      'Source:    ' + (p.source  || '—') + '\n' +
      'Page:      ' + (p.page    || '—') + '\n' +
      'Time:      ' + ts + '\n';

    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: subject,
      body: body,
      replyTo: p.email || NOTIFY_EMAIL
    });

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('TREVOC lead endpoint is live.');
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
