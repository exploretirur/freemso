/**
 * AI Driven MS Office — Enquiry receiver  (v2 — debuggable)
 * Paste this into Extensions > Apps Script in your Google Sheet.
 */

/* ============================================================
   STEP 1: Paste your Spreadsheet ID between the quotes below.

   Find it in your Sheet's address bar:
   https://docs.google.com/spreadsheets/d/THIS_LONG_PART_HERE/edit
                                          ^^^^^^^^^^^^^^^^^^
   Leaving it blank works ONLY if this script was opened from
   the Sheet's Extensions > Apps Script menu.
   ============================================================ */
var SHEET_ID = "";

var HEADERS = [
  'Enquiry ID', 'Date & Time', 'Name', 'C/o - S/o - D/o', 'Address', 'Place', 'District',
  'Mobile', 'Email', 'Qualification', 'Group / Course', 'School / College'
];


/** Accepts both GET (?name=...) and POST (JSON body). */
function doGet(e)  { return handle(e); }
function doPost(e) { return handle(e); }


function handle(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    var d = readData(e);

    // No data at all = someone just opened the URL in a browser.
    if (!d || (!d.name && !d.mobile && !d.enquiryId)) {
      return reply({
        result: 'ok',
        message: 'Enquiry endpoint is running. Sheet connection: ' + testSheet()
      });
    }

    var sheet = getSheet();

    sheet.appendRow([
      d.enquiryId     || '',
      d.timestamp     || new Date().toLocaleString(),
      d.name          || '',
      d.guardian      || '',
      d.address       || '',
      d.place         || '',
      d.district      || '',
      "'" + (d.mobile || ''),   // leading quote preserves leading digits
      d.email         || '',
      d.qualification || '',
      d.courseGroup   || '',
      d.institution   || ''
    ]);

    return reply({ result: 'success', id: d.enquiryId });

  } catch (err) {
    return reply({ result: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}


/** Pulls the record out of whichever way it arrived. */
function readData(e) {
  if (!e) return null;

  // POST with a JSON body
  if (e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (err) { /* fall through */ }
  }
  // GET or POST with query / form parameters
  if (e.parameter && Object.keys(e.parameter).length) return e.parameter;

  return null;
}


function getSheet() {
  var ss = SHEET_ID
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'No spreadsheet found. This script is not attached to a Sheet — ' +
      'set SHEET_ID at the top of the script.'
    );
  }

  var sheet = ss.getSheetByName('Enquiries');
  if (!sheet) {
    sheet = ss.insertSheet('Enquiries');
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
         .setFontWeight('bold')
         .setBackground('#2454C7')
         .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 130);
    sheet.setColumnWidth(2, 150);
  }
  return sheet;
}


function testSheet() {
  try {
    var s = getSheet();
    return 'OK - writing to "' + s.getParent().getName() + '" (' +
           (s.getLastRow() - 1) + ' enquiries so far)';
  } catch (err) {
    return 'FAILED - ' + err.toString();
  }
}


function reply(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* ============================================================
   Run this from the Apps Script editor (select runTest, press Run)
   to confirm the script can write to your Sheet before you even
   touch the website. Then check the Execution log.
   ============================================================ */
function runTest() {
  var sheet = getSheet();
  sheet.appendRow([
    'TEST-0001', new Date().toLocaleString(), 'Test Student', 'S/o Test Guardian', 'Test Address',
    'Tirur', 'Malappuram', "'9999999999", 'test@example.com', 'Degree', 'B.Com', 'Test College'
  ]);
  Logger.log('SUCCESS - test row added to: ' + sheet.getParent().getName());
}
