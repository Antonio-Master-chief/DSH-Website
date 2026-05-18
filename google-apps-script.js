// ================================================================
// DSH Institute of Technology — Google Apps Script Web App
// ================================================================
// HOW TO DEPLOY:
//   1. Go to script.google.com → open the DSH project
//   2. Select all and replace with this file
//   3. Fill in DRIVE_FOLDER_ID below if you want file uploads
//   4. Deploy → Manage Deployments → Edit → New version → Deploy
// ================================================================

var CONFIG = {
  ENROLLMENT_SPREADSHEET_ID:    '1maYMinVkK_dWTTX94TdlNtXnlJhe_QJjJviOZBJUSRg',
  ENQUIRY_SPREADSHEET_ID:       '1hAEwhrhJZ8U5LhTt0NYcx4hMvP2e5-0EuNDAc2zHE14',
  ACCOMMODATION_SPREADSHEET_ID: '16yUQtnER39njj4L6xF62iCNbROkN7m7D4Q4ypqBn28U',
  DRIVE_FOLDER_ID:              'PASTE_YOUR_DRIVE_FOLDER_ID_HERE',
  WA_NUMBER:                    '60132831908',
  WA_API_KEY:                   'PASTE_YOUR_CALLMEBOT_API_KEY_HERE',
  ADMIN_EMAIL:                  'admissions@dit.edu.my'
};

// ── Entry points ────────────────────────────────────────────────

function doGet(e) {
  if (e && e.parameter && e.parameter.data) {
    try {
      var data = JSON.parse(e.parameter.data);
      if      (data.type === 'enrollment')     processEnrollment(data);
      else if (data.type === 'enquiry')        processEnquiry(data);
      else if (data.type === 'accommodation')  processAccommodation(data);
    } catch (err) {
      Logger.log('doGet error: ' + err.toString());
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if      (data.type === 'enrollment')    processEnrollment(data);
    else if (data.type === 'enquiry')       processEnquiry(data);
    else if (data.type === 'accommodation') processAccommodation(data);
    else throw new Error('Unknown type: ' + data.type);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Sheet helpers ───────────────────────────────────────────────

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#1a1d2e').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

var ENR_HEADERS = [
  'Timestamp','Full Name','Email','Phone','ID Type','ID Number',
  'Date of Birth','Nationality','Programme','Education Level',
  'Heard About Us','Message','Student Type',
  'Health Declaration (Drive)','Registration Form (Drive)',
  'Status','Admin Notes','Date Checked'
];
var ENQ_HEADERS = [
  'Timestamp','Full Name','Email','Phone','Message',
  'Status','Admin Notes','Date Checked'
];
var ACM_HEADERS = [
  'Timestamp','Full Name','Email','Phone','Gender',
  'Student Type','ID Number','Programme','Move-in Date',
  'Status','Admin Notes','Date Checked'
];

// ── Enrollment handler ──────────────────────────────────────────

function processEnrollment(data) {
  var ss    = SpreadsheetApp.openById(CONFIG.ENROLLMENT_SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, 'Enrollments', ENR_HEADERS);
  var ts    = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'dd/MM/yyyy HH:mm:ss');

  var safeName   = (data.fullName || 'Student').replace(/[^A-Za-z0-9 ]/g, '').replace(/\s+/g, '_');
  var healthLink = '';
  var regLink    = '';

  var hasDriveFolder = CONFIG.DRIVE_FOLDER_ID && CONFIG.DRIVE_FOLDER_ID.indexOf('PASTE') === -1;
  if (hasDriveFolder) {
    var folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    if (data.healthFile && data.healthFile.content) {
      var hBlob = Utilities.newBlob(
        Utilities.base64Decode(data.healthFile.content),
        data.healthFile.mimeType,
        safeName + '_HealthDeclaration.' + data.healthFile.ext
      );
      healthLink = folder.createFile(hBlob).getUrl();
    }
    if (data.regFile && data.regFile.content) {
      var rBlob = Utilities.newBlob(
        Utilities.base64Decode(data.regFile.content),
        data.regFile.mimeType,
        safeName + '_RegistrationForm.' + data.regFile.ext
      );
      regLink = folder.createFile(rBlob).getUrl();
    }
  }

  sheet.appendRow([
    ts,
    data.fullName    || '',
    data.email       || '',
    data.phone       || '',
    (data.idType || '').toUpperCase(),
    data.idNumber    || '',
    data.dob         || '',
    data.nationality || '',
    data.programme   || '',
    data.education   || '',
    data.source      || 'Not specified',
    data.message     || '',
    data.nationality === 'Malaysian' ? 'Local' : 'International',
    healthLink,
    regLink,
    'Pending', '', ''
  ]);

  var lastRow  = sheet.getLastRow();
  styleStatusCell(sheet, lastRow, 16);

  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + CONFIG.ENROLLMENT_SPREADSHEET_ID;

  sendWhatsApp(
    '📋 *NEW ENROLMENT — DIT*\n────────────────────\n' +
    '👤 ' + (data.fullName   || '') + '\n' +
    '📧 ' + (data.email      || '') + '\n' +
    '📱 ' + (data.phone      || '') + '\n' +
    '🌏 ' + (data.nationality|| '') + '\n' +
    '🎓 ' + (data.programme  || '') + '\n' +
    '📅 ' + ts + '\n────────────────────\n📊 ' + sheetUrl
  );

  sendEmail(
    '📋 New Enrolment: ' + (data.fullName || 'Student') + ' — DIT',
    buildEmailHtml('📋 New Enrolment — DIT', [
      ['Full Name',      data.fullName    || ''],
      ['Email',          data.email       || ''],
      ['Phone',          data.phone       || ''],
      ['Nationality',    data.nationality || ''],
      ['Programme',      data.programme   || ''],
      ['Education Level',data.education   || ''],
      ['ID Type',        (data.idType || '').toUpperCase()],
      ['ID Number',      data.idNumber    || ''],
      ['Date of Birth',  data.dob         || ''],
      ['Heard About Us', data.source      || ''],
      ['Message',        data.message     || ''],
      ['Submitted',      ts]
    ], sheetUrl)
  );
}

// ── Enquiry handler ─────────────────────────────────────────────

function processEnquiry(data) {
  var ss    = SpreadsheetApp.openById(CONFIG.ENQUIRY_SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, 'Enquiries', ENQ_HEADERS);
  var ts    = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'dd/MM/yyyy HH:mm:ss');

  sheet.appendRow([
    ts,
    data.name    || '',
    data.email   || '',
    data.phone   || '',
    data.message || '',
    'Pending', '', ''
  ]);

  var lastRow  = sheet.getLastRow();
  styleStatusCell(sheet, lastRow, 6);

  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + CONFIG.ENQUIRY_SPREADSHEET_ID;

  sendWhatsApp(
    '💬 *NEW ENQUIRY — DIT*\n────────────────────\n' +
    '👤 ' + (data.name    || '') + '\n' +
    '📧 ' + (data.email   || '') + '\n' +
    '📱 ' + (data.phone   || '') + '\n' +
    '💬 ' + (data.message || '').substring(0, 250) + '\n' +
    '📅 ' + ts + '\n────────────────────\n📊 ' + sheetUrl
  );

  sendEmail(
    '💬 New Enquiry: ' + (data.name || 'Visitor') + ' — DIT',
    buildEmailHtml('💬 New Enquiry — DIT', [
      ['Full Name', data.name    || ''],
      ['Email',     data.email   || ''],
      ['Phone',     data.phone   || ''],
      ['Message',   data.message || ''],
      ['Submitted', ts]
    ], sheetUrl)
  );
}

// ── Accommodation handler ────────────────────────────────────────

function processAccommodation(data) {
  var ss    = SpreadsheetApp.openById(CONFIG.ACCOMMODATION_SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, 'Accommodation', ACM_HEADERS);
  var ts    = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'dd/MM/yyyy HH:mm:ss');

  sheet.appendRow([
    ts,
    data.fullName    || '',
    data.email       || '',
    data.phone       || '',
    data.gender      || '',
    data.studentType || '',
    data.idNumber    || '',
    data.programme   || '',
    data.moveIn      || '',
    'Pending', '', ''
  ]);

  var lastRow  = sheet.getLastRow();
  styleStatusCell(sheet, lastRow, 10);

  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + CONFIG.ACCOMMODATION_SPREADSHEET_ID;

  sendWhatsApp(
    '🏠 *NEW ACCOMMODATION — DIT*\n────────────────────\n' +
    '👤 ' + (data.fullName  || '') + '\n' +
    '📧 ' + (data.email     || '') + '\n' +
    '📱 ' + (data.phone     || '') + '\n' +
    '🎓 ' + (data.programme || '') + '\n' +
    '📅 Move-in: ' + (data.moveIn || '') + '\n' +
    '────────────────────\n📊 ' + sheetUrl
  );

  sendEmail(
    '🏠 New Accommodation Application: ' + (data.fullName || 'Student') + ' — DIT',
    buildEmailHtml('🏠 New Accommodation Application — DIT', [
      ['Full Name',    data.fullName    || ''],
      ['Email',        data.email       || ''],
      ['Phone',        data.phone       || ''],
      ['Gender',       data.gender      || ''],
      ['Student Type', data.studentType || ''],
      ['ID Number',    data.idNumber    || ''],
      ['Programme',    data.programme   || ''],
      ['Move-in Date', data.moveIn      || ''],
      ['Submitted',    ts]
    ], sheetUrl)
  );
}

// ── Email builder ───────────────────────────────────────────────

function buildEmailHtml(title, rows, sheetUrl) {
  var rowsHtml = rows.map(function(r) {
    return '<tr><td style="padding:8px 12px 8px 0;color:#555;width:40%;vertical-align:top"><b>' +
           r[0] + '</b></td><td style="padding:8px 0">' + r[1] + '</td></tr>';
  }).join('');
  return '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;border:1px solid #ddd;border-radius:8px;overflow:hidden">' +
    '<div style="background:#1a2063;padding:20px 24px">' +
    '<h2 style="color:#fff;margin:0;font-size:18px">' + title + '</h2></div>' +
    '<div style="padding:24px;background:#fff">' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px">' + rowsHtml + '</table>' +
    '<div style="margin-top:20px">' +
    '<a href="' + sheetUrl + '" style="background:#1a2063;color:#fff;padding:10px 20px;border-radius:5px;text-decoration:none;font-size:14px">View Spreadsheet</a>' +
    '</div></div></div>';
}

// ── WhatsApp via CallMeBot ──────────────────────────────────────

function sendWhatsApp(message) {
  try {
    if (!CONFIG.WA_API_KEY || CONFIG.WA_API_KEY.indexOf('PASTE') !== -1) return;
    var url = 'https://api.callmebot.com/whatsapp.php' +
      '?phone='  + encodeURIComponent(CONFIG.WA_NUMBER) +
      '&text='   + encodeURIComponent(message) +
      '&apikey=' + encodeURIComponent(CONFIG.WA_API_KEY);
    UrlFetchApp.fetch(url);
  } catch (e) {
    Logger.log('WhatsApp error: ' + e.toString());
  }
}

// ── Email via Gmail ─────────────────────────────────────────────

function sendEmail(subject, htmlBody) {
  try {
    MailApp.sendEmail({ to: CONFIG.ADMIN_EMAIL, subject: subject, htmlBody: htmlBody });
  } catch (e) {
    Logger.log('Email error: ' + e.toString());
  }
}

// ── Status cell styling ─────────────────────────────────────────

function styleStatusCell(sheet, row, col) {
  var cell = sheet.getRange(row, col);
  cell.setBackground('#FFF3CD');
  cell.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending','In Review','Contacted','Done','Rejected'], true)
      .build()
  );
}

// ── One-time setup functions — run each ONCE from the script editor ─

function setupEnrollmentSheet() {
  var ss  = SpreadsheetApp.openById(CONFIG.ENROLLMENT_SPREADSHEET_ID);
  var sh  = ss.getSheetByName('Enrollments') || ss.insertSheet('Enrollments');
  sh.clearContents();
  sh.getRange(1,1,1,ENR_HEADERS.length).setValues([ENR_HEADERS])
    .setBackground('#1a1d2e').setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);
  // col: 1=Timestamp, 2=Full Name, 3=Email, 4=Phone, 5=ID Type, 6=ID Number,
  //      7=DOB, 8=Nationality, 9=Programme, 10=Edu Level, 11=Heard About Us,
  //      12=Message, 13=Student Type, 14=Health Dec, 15=Reg Form, 16=Status,
  //      17=Admin Notes, 18=Date Checked
  var widths = [1,150, 2,180, 3,210, 4,120, 5,80, 6,140,
                7,110, 8,110, 9,270, 10,140, 11,160,
                12,270, 13,100, 14,220, 15,220,
                16,100, 17,230, 18,120];
  for (var i = 0; i < widths.length; i += 2) sh.setColumnWidth(widths[i], widths[i+1]);
  Logger.log('Enrollment sheet ready!');
}

function setupEnquirySheet() {
  var ss  = SpreadsheetApp.openById(CONFIG.ENQUIRY_SPREADSHEET_ID);
  var sh  = ss.getSheetByName('Enquiries') || ss.insertSheet('Enquiries');
  sh.clearContents();
  sh.getRange(1,1,1,ENQ_HEADERS.length).setValues([ENQ_HEADERS])
    .setBackground('#1a1d2e').setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);
  // col: 1=Timestamp, 2=Full Name, 3=Email, 4=Phone, 5=Message,
  //      6=Status, 7=Admin Notes, 8=Date Checked
  var widths = [1,150, 2,180, 3,210, 4,120, 5,350, 6,100, 7,230, 8,120];
  for (var i = 0; i < widths.length; i += 2) sh.setColumnWidth(widths[i], widths[i+1]);
  Logger.log('Enquiry sheet ready!');
}

function setupAccommodationSheet() {
  var ss  = SpreadsheetApp.openById(CONFIG.ACCOMMODATION_SPREADSHEET_ID);
  var sh  = ss.getSheetByName('Accommodation') || ss.insertSheet('Accommodation');
  sh.clearContents();
  sh.getRange(1,1,1,ACM_HEADERS.length).setValues([ACM_HEADERS])
    .setBackground('#1a1d2e').setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);
  // col: 1=Timestamp, 2=Full Name, 3=Email, 4=Phone, 5=Gender,
  //      6=Student Type, 7=ID Number, 8=Programme, 9=Move-in Date,
  //      10=Status, 11=Admin Notes, 12=Date Checked
  var widths = [1,150, 2,180, 3,210, 4,120, 5,80,
                6,100, 7,140, 8,270, 9,120,
                10,100, 11,230, 12,120];
  for (var i = 0; i < widths.length; i += 2) sh.setColumnWidth(widths[i], widths[i+1]);
  Logger.log('Accommodation sheet ready!');
}
