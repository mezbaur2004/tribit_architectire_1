/**
 * Tribid call-back form -> Google Sheet + email notification.
 *
 * Setup:
 *  1. Create a Google Sheet, copy its ID from the URL:
 *       https://docs.google.com/spreadsheets/d/<THIS_PART_IS_THE_ID>/edit
 *     and paste it into SHEET_ID below.
 *  2. (Optional) set NOTIFY_EMAIL. Leave it empty to email the account that owns this script.
 *  3. Run testSubmit() once from the editor and approve the permissions
 *     (Sheets + sending mail). Check the sheet and your inbox.
 *  4. Deploy > New deployment > type "Web app"
 *       Execute as: Me
 *       Who has access: Anyone
 *     Copy the web app URL (ends in /exec) into APPS_SCRIPT_URL in index.html.
 *  5. After ANY change to this script, use Deploy > Manage deployments > Edit >
 *     Version: New version, otherwise the live URL keeps running the old code.
 */

var SHEET_ID = 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE';
var SHEET_NAME = 'Leads';   // tab name; created automatically if missing
var NOTIFY_EMAIL = '';      // e.g. 'you@example.com'; empty = the script owner's email
var TIMEZONE = 'Asia/Dhaka';

var HEADERS = ['Received at', 'Phone', 'Service', 'Submitted from browser at'];

function doPost(e) {
  try {
    var p = (e && e.parameter) || {};
    var phone = normalizePhone_(p.phone);
    if (!phone) return json_({ ok: false, error: 'invalid phone' });

    var service = cleanText_(p.service, 120) || 'Not specified';
    var submittedAt = cleanText_(p.submittedAt, 40);

    // One writer at a time so simultaneous submissions never overwrite each other.
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var sheet = getSheet_();
      var row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1, 1, 4).setNumberFormat('@'); // keep the leading 0 of the phone number
      sheet.getRange(row, 1, 1, 4).setValues([[
        Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
        phone,
        service,
        submittedAt
      ]]);
    } finally {
      lock.releaseLock();
    }

    notify_(phone, service);
    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  }
}

// Opening the /exec URL in a browser shows this, handy for checking the deployment.
function doGet() {
  return json_({ ok: true, message: 'Tribid form endpoint is running' });
}

/* ---------- helpers ---------- */

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function notify_(phone, service) {
  var to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
  var when = Utilities.formatDate(new Date(), TIMEZONE, 'dd MMM yyyy, hh:mm a');
  MailApp.sendEmail({
    to: to,
    subject: 'New registration: ' + phone + ' (' + service + ')',
    body: 'Someone registered with the phone number ' + phone + '.\n\n' +
          'Service: ' + service + '\n' +
          'Time: ' + when + '\n\n' +
          'Sheet: https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit'
  });
}

// Bangladeshi mobile: 01XXXXXXXXX, optionally prefixed by +88 / 88. Returns 01XXXXXXXXX or null.
function normalizePhone_(raw) {
  var d = String(raw || '').replace(/[\s\-()]/g, '');
  if (/^\+?880?1[3-9]\d{8}$/.test(d)) d = '0' + d.replace(/^\+?880?/, '');
  return /^01[3-9]\d{8}$/.test(d) ? d : null;
}

// Trim and cap the length. Cells are formatted as plain text, so "=..." can never run as a formula.
function cleanText_(s, max) {
  return String(s || '').replace(/[\r\n\t]+/g, ' ').trim().substring(0, max);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- run this once from the editor to authorize and test ---------- */
function testSubmit() {
  var res = doPost({ parameter: { phone: '01712345678', service: 'Test', submittedAt: new Date().toISOString() } });
  Logger.log(res.getContent());
}
