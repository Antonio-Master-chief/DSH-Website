// ================================================================
// DSH Institute of Technology — Google Apps Script Web App
// ================================================================
// HOW TO DEPLOY:
//   1. Go to script.google.com → New Project → paste this whole file
//   2. Fill in the 4 CONFIG values below
//   3. Run setupSheets() once (from the editor) to create the sheets
//   4. Click Deploy → New Deployment → Web App
//      Execute as: Me | Who has access: Anyone
//   5. Copy the Web App URL and paste it into enroll.html
// ================================================================

var CONFIG = {
  SPREADSHEET_ID: 'PASTE_YOUR_SPREADSHEET_ID_HERE',
  DRIVE_FOLDER_ID: 'PASTE_YOUR_DRIVE_FOLDER_ID_HERE',
  WA_NUMBER:       '60132831908',
  WA_API_KEY:      'PASTE_YOUR_CALLMEBOT_API_KEY_HERE'
};

// ── Entry points ────────────────────────────────────────────────

function doGet(e) {
  if (e && e.parameter && e.parameter.data) {
    try {
      var data = JSON.parse(e.parameter.data);
      if (data.type === 'enrollment')  processEnrollment(data);
      else if (data.type === 'enquiry') processEnquiry(data);
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
    if (data.type === 'enrollment') {
      processEnrollment(data);
    } else if (data.type === 'enquiry') {
      processEnquiry(data);
    } else {
      throw new Error('Unknown type: ' + data.type);
    }
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

// ── Enrollment handler ──────────────────────────────────────────

function processEnrollment(data) {
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, 'Enrollments', ENR_HEADERS);
  var ts     = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'dd/MM/yyyy HH:mm:ss');
  var folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  var safeName = (data.fullName || 'Student').replace(/[^A-Za-z0-9 ]/g, '').replace(/\s+/g, '_');

  var healthLink = '';
  var regLink    = '';

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

  sheet.appendRow([
    ts,
    data.fullName      || '',
    data.email         || '',
    data.phone         || '',
    (data.idType || '').toUpperCase(),
    data.idNumber      || '',
    data.dob           || '',
    data.nationality   || '',
    data.programme     || '',
    data.education     || '',
    data.source        || 'Not specified',
    data.message       || '',
    data.nationality === 'Malaysian' ? 'Local' : 'International',
    healthLink,
    regLink,
    'Pending',  // Status  ← admin fills this
    '',         // Notes   ← admin fills this
    ''          // Date Checked ← admin fills this
  ]);

  var lastRow = sheet.getLastRow();
  styleStatusCell(sheet, lastRow, 16);  // col P = Status

  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + CONFIG.SPREADSHEET_ID;
  var waMsg =
    '📋 *NEW ENROLMENT — DIT*\n' +
    '────────────────────\n' +
    '👤 ' + (data.fullName || '') + '\n' +
    '📧 ' + (data.email    || '') + '\n' +
    '📱 ' + (data.phone    || '') + '\n' +
    '🌏 ' + (data.nationality || '') + '\n' +
    '🎓 ' + (data.programme   || '') + '\n' +
    '📅 ' + ts + '\n' +
    '────────────────────\n' +
    '📊 View sheet:\n' + sheetUrl;

  sendWhatsApp(waMsg);
}

// ── Enquiry handler ─────────────────────────────────────────────

function processEnquiry(data) {
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, 'Enquiries', ENQ_HEADERS);
  var ts    = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'dd/MM/yyyy HH:mm:ss');

  sheet.appendRow([
    ts,
    data.name    || '',
    data.email   || '',
    data.phone   || '',
    data.message || '',
    'Pending',  // Status  ← admin fills this
    '',         // Notes   ← admin fills this
    ''          // Date Checked ← admin fills this
  ]);

  var lastRow = sheet.getLastRow();
  styleStatusCell(sheet, lastRow, 6);  // col F = Status

  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + CONFIG.SPREADSHEET_ID;
  var waMsg =
    '💬 *NEW ENQUIRY — DIT*\n' +
    '────────────────────\n' +
    '👤 ' + (data.name  || '') + '\n' +
    '📧 ' + (data.email || '') + '\n' +
    '📱 ' + (data.phone || '') + '\n' +
    '💬 ' + (data.message || '').substring(0, 250) + '\n' +
    '📅 ' + ts + '\n' +
    '────────────────────\n' +
    '📊 View sheet:\n' + sheetUrl;

  sendWhatsApp(waMsg);
}

// ── WhatsApp via CallMeBot ──────────────────────────────────────

function sendWhatsApp(message) {
  try {
    var url = 'https://api.callmebot.com/whatsapp.php' +
      '?phone='  + encodeURIComponent(CONFIG.WA_NUMBER) +
      '&text='   + encodeURIComponent(message) +
      '&apikey=' + encodeURIComponent(CONFIG.WA_API_KEY);
    UrlFetchApp.fetch(url);
  } catch (e) {
    Logger.log('WhatsApp error: ' + e.toString());
  }
}

// ── Helpers ─────────────────────────────────────────────────────

function styleStatusCell(sheet, row, col) {
  var cell = sheet.getRange(row, col);
  cell.setBackground('#FFF3CD');
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'In Review', 'Contacted', 'Done', 'Rejected'], true)
    .build();
  cell.setDataValidation(rule);
}

// ── One-time sheet setup — run this ONCE from the script editor ─
// Select this function in the dropdown and click Run.

function setupSheets() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Enrollments
  var enr = ss.getSheetByName('Enrollments');
  if (!enr) { enr = ss.insertSheet('Enrollments'); }
  enr.clearContents();
  var enrH = [
    'Timestamp', 'Full Name', 'Email', 'Phone',
    'ID Type', 'ID Number', 'Date of Birth', 'Nationality',
    'Programme', 'Education Level', 'Heard About Us', 'Message',
    'Student Type', 'Health Declaration (Drive)', 'Registration Form (Drive)',
    'Status', 'Admin Notes', 'Date Checked'
  ];
  enr.getRange(1, 1, 1, enrH.length).setValues([enrH])
    .setBackground('#1a1d2e').setFontColor('#ffffff').setFontWeight('bold');
  enr.setFrozenRows(1);
  [1,140, 2,160, 3,190, 4,120, 5,90, 6,130, 7,110, 8,110,
   9,230, 10,140, 11,160, 12,220, 13,100, 14,220, 15,220,
   16,100, 17,220, 18,120].forEach(function(v,i,a){
    if (i%2===0) enr.setColumnWidth(a[i], a[i+1]);
  });

  // Enquiries
  var enq = ss.getSheetByName('Enquiries');
  if (!enq) { enq = ss.insertSheet('Enquiries'); }
  enq.clearContents();
  var enqH = ['Timestamp', 'Full Name', 'Email', 'Phone', 'Message',
               'Status', 'Admin Notes', 'Date Checked'];
  enq.getRange(1, 1, 1, enqH.length).setValues([enqH])
    .setBackground('#1a1d2e').setFontColor('#ffffff').setFontWeight('bold');
  enq.setFrozenRows(1);
  [1,140, 2,160, 3,190, 4,120, 5,320, 6,100, 7,220, 8,120].forEach(function(v,i,a){
    if (i%2===0) enq.setColumnWidth(a[i], a[i+1]);
  });

  Logger.log('Sheets set up! Enrollments and Enquiries are ready.');
}
