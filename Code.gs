// ── Super Grocer Pharmacy – Form Handler ──────────────────────────────────
// Deploy: Extensions → Apps Script → Deploy → New deployment
//         Type: Web app | Execute as: Me | Who has access: Anyone
// ─────────────────────────────────────────────────────────────────────────

var SPREADSHEET_ID  = '1cburRYkFqfup3dgX7MnXTWZZC4aFu4u0CLonMkCdRaU';
var DRIVE_FOLDER_ID = '1jKXu1O2wz0Ku8r5cwplnCUEQ-y1aBvxo';
var NOTIFY_EMAIL    = 'tracy.jia1223@gmail.com';

// Tab GIDs — stable even if tabs are renamed
var SHEET_GID = {
  contact  : 0,
  transfer : 1671050477,
  refill   : 1300104817
};

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    var ts   = Utilities.formatDate(new Date(), 'America/Vancouver', 'yyyy-MM-dd HH:mm:ss');

    // ── File upload to Drive ──────────────────────────────────────────────
    var attachUrl = '';
    if (data.fileData && data.fileName) {
      var folder  = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      var decoded = Utilities.base64Decode(data.fileData);
      var blob    = Utilities.newBlob(decoded, data.fileType || 'application/octet-stream', data.fileName);
      var file    = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      attachUrl = file.getUrl();
    }

    // ── Route to correct sheet & build email ─────────────────────────────
    var sheet, row, subject, body;

    if (data.formType === 'refill') {
      sheet = ss.getSheetById(SHEET_GID.refill);
      row     = [ts, data.firstName, data.lastName, data.phone, data.email || '', data.prescription, attachUrl];
      subject = '[Super Grocer Pharmacy] Refill Request – ' + data.firstName + ' ' + data.lastName;
      body    = fmt('Refill Prescription Request', {
        'Name'        : data.firstName + ' ' + data.lastName,
        'Phone'       : data.phone,
        'Email'       : data.email || '—',
        'Prescription': data.prescription,
        'Attachment'  : attachUrl || 'None'
      }, ts);

    } else if (data.formType === 'transfer') {
      sheet = ss.getSheetById(SHEET_GID.transfer);
      row     = [ts, data.firstName, data.lastName, data.phone, data.email || '', data.currentPharmacy, data.prescription, attachUrl];
      subject = '[Super Grocer Pharmacy] Transfer Request – ' + data.firstName + ' ' + data.lastName;
      body    = fmt('Transfer Prescription Request', {
        'Name'            : data.firstName + ' ' + data.lastName,
        'Phone'           : data.phone,
        'Email'           : data.email || '—',
        'Current Pharmacy': data.currentPharmacy,
        'Prescription'    : data.prescription,
        'Attachment'      : attachUrl || 'None'
      }, ts);

    } else if (data.formType === 'contact') {
      sheet = ss.getSheetById(SHEET_GID.contact);
      row     = [ts, data.firstName, data.lastName, data.phone || '', data.email, data.message];
      subject = '[Super Grocer Pharmacy] Contact Message – ' + data.firstName + ' ' + data.lastName;
      body    = fmt('Contact Form', {
        'Name'   : data.firstName + ' ' + data.lastName,
        'Phone'  : data.phone || '—',
        'Email'  : data.email,
        'Message': data.message
      }, ts);

    } else {
      throw new Error('Unknown formType: ' + data.formType);
    }

    sheet.appendRow(row);
    GmailApp.sendEmail(NOTIFY_EMAIL, subject, body);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log(err);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────

function fmt(title, fields, ts) {
  var lines = title + '\n' + Array(42).join('─') + '\n\n';
  for (var k in fields) lines += k + ': ' + fields[k] + '\n';
  lines += '\nReceived : ' + ts + ' (Pacific Time)';
  lines += '\nSource   : Super Grocer Pharmacy Website';
  return lines;
}
