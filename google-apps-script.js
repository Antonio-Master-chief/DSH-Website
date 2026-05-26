// ================================================================
// DSH Institute of Technology — Google Apps Script Web App
// ================================================================

var CONFIG = {
  ENROLLMENT_SPREADSHEET_ID:    '171gtxedQgXIgyEALSHXZA6YX2NI_PBs2zFQ4tL97nQs',
  ENQUIRY_SPREADSHEET_ID:       '1lco16uCbAcn5b8FVTzlsIXG_td6rxoRjZ1Fs9DBfgWM',
  ACCOMMODATION_SPREADSHEET_ID: '1tQ9cVBft-DQIZwZMD1RqSyCVGypxiHCmLMRTVKdwCF8',
  DRIVE_FOLDER_ID:              '15GshPwka4TJQnc6NLpCA5vlAcpc7v55q',
  WA_NUMBER:                    '60132831908',
  WA_API_KEY:                   'PASTE_YOUR_CALLMEBOT_API_KEY_HERE',
  ADMIN_EMAIL:                  'admissions@dit.edu.my'
};

// ── Entry points ────────────────────────────────────────────────

function doGet(e) {
  if (e && e.parameter && e.parameter.data) {
    try {
      var data = JSON.parse(e.parameter.data);
      if      (data.type === 'enrollment')    processEnrollment(data);
      else if (data.type === 'enquiry')       processEnquiry(data);
      else if (data.type === 'accommodation') processAccommodation(data);
    } catch (err) { Logger.log('doGet error: ' + err.toString()); }
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

// ── Headers ─────────────────────────────────────────────────────

var ENR_HEADERS = [
  '#','Timestamp','Full Name','Email','Phone','ID Type','ID Number',
  'Date of Birth','Nationality','Programme','Education Level',
  'Heard About Us','Message','Student Type',
  'Health Declaration (Drive)','Registration Form (Drive)',
  'Status','Admin Notes','Date Checked'
];
var ENQ_HEADERS = [
  '#','Timestamp','Full Name','Email','Phone',
  'Message','Status','Admin Notes','Date Checked'
];
var ACM_HEADERS = [
  '#','Timestamp','Full Name','Email','Phone','Gender',
  'Student Type','ID Number','Programme','Move-in Date',
  'Status','Admin Notes','Date Checked'
];

// ── Sheet helper ─────────────────────────────────────────────────

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

// ── Enrollment handler ──────────────────────────────────────────

function processEnrollment(data) {
  var ss    = SpreadsheetApp.openById(CONFIG.ENROLLMENT_SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, 'Enrollments', ENR_HEADERS);
  var ts    = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'dd/MM/yyyy HH:mm:ss');
  var rowNum = sheet.getLastRow(); // row number = rows already present (header counts as 1)

  var safeName   = (data.fullName || 'Student').replace(/[^A-Za-z0-9 ]/g, '').replace(/\s+/g, '_');
  var healthLink = '';
  var regLink    = '';
  var hasDrive   = CONFIG.DRIVE_FOLDER_ID && CONFIG.DRIVE_FOLDER_ID.indexOf('PASTE') === -1;

  if (hasDrive) {
    var folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    if (data.healthFile && data.healthFile.content) {
      healthLink = folder.createFile(Utilities.newBlob(
        Utilities.base64Decode(data.healthFile.content),
        data.healthFile.mimeType,
        safeName + '_HealthDeclaration.' + data.healthFile.ext
      )).getUrl();
    }
    if (data.regFile && data.regFile.content) {
      regLink = folder.createFile(Utilities.newBlob(
        Utilities.base64Decode(data.regFile.content),
        data.regFile.mimeType,
        safeName + '_RegistrationForm.' + data.regFile.ext
      )).getUrl();
    }
  }

  sheet.appendRow([
    rowNum,
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

  styleStatusCell(sheet, sheet.getLastRow(), 17); // col Q = Status

  var sheetUrl     = 'https://docs.google.com/spreadsheets/d/' + CONFIG.ENROLLMENT_SPREADSHEET_ID;
  var driveFolderUrl = hasDrive ? 'https://drive.google.com/drive/folders/' + CONFIG.DRIVE_FOLDER_ID : '';
  sendWhatsApp('📋 *NEW ENROLMENT — DIT*\n────────────────────\n' +
    '👤 ' + (data.fullName || '') + '\n📧 ' + (data.email || '') + '\n📱 ' + (data.phone || '') +
    '\n🌏 ' + (data.nationality || '') + '\n🎓 ' + (data.programme || '') + '\n📅 ' + ts +
    '\n────────────────────\n📊 ' + sheetUrl +
    (driveFolderUrl ? '\n📁 Uploaded Files: ' + driveFolderUrl : ''));
  sendEmail('📋 New Enrolment: ' + (data.fullName || 'Student') + ' — DIT',
    buildEmailHtml('📋 New Enrolment — DIT', [
      ['Full Name',       data.fullName    || ''],
      ['Email',           data.email       || ''],
      ['Phone',           data.phone       || ''],
      ['Nationality',     data.nationality || ''],
      ['Programme',       data.programme   || ''],
      ['Education Level', data.education   || ''],
      ['ID Type',         (data.idType || '').toUpperCase()],
      ['ID Number',       data.idNumber    || ''],
      ['Date of Birth',   data.dob         || ''],
      ['Heard About Us',  data.source      || ''],
      ['Message',         data.message     || ''],
      ['Submitted',       ts]
    ], sheetUrl, driveFolderUrl));
}

// ── Enquiry handler ─────────────────────────────────────────────

function processEnquiry(data) {
  var ss     = SpreadsheetApp.openById(CONFIG.ENQUIRY_SPREADSHEET_ID);
  var sheet  = getOrCreateSheet(ss, 'Enquiries', ENQ_HEADERS);
  var ts     = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'dd/MM/yyyy HH:mm:ss');
  var rowNum = sheet.getLastRow();

  sheet.appendRow([
    rowNum,
    ts,
    data.name    || '',
    data.email   || '',
    data.phone   || '',
    data.message || '',
    'Pending', '', ''
  ]);

  styleStatusCell(sheet, sheet.getLastRow(), 7); // col G = Status

  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + CONFIG.ENQUIRY_SPREADSHEET_ID;
  sendWhatsApp('💬 *NEW ENQUIRY — DIT*\n────────────────────\n' +
    '👤 ' + (data.name || '') + '\n📧 ' + (data.email || '') + '\n📱 ' + (data.phone || '') +
    '\n💬 ' + (data.message || '').substring(0, 250) + '\n📅 ' + ts +
    '\n────────────────────\n📊 ' + sheetUrl);
  sendEmail('💬 New Enquiry: ' + (data.name || 'Visitor') + ' — DIT',
    buildEmailHtml('💬 New Enquiry — DIT', [
      ['Full Name', data.name    || ''],
      ['Email',     data.email   || ''],
      ['Phone',     data.phone   || ''],
      ['Message',   data.message || ''],
      ['Submitted', ts]
    ], sheetUrl));
}

// ── Accommodation handler ────────────────────────────────────────

function processAccommodation(data) {
  var ss     = SpreadsheetApp.openById(CONFIG.ACCOMMODATION_SPREADSHEET_ID);
  var sheet  = getOrCreateSheet(ss, 'Accommodation', ACM_HEADERS);
  var ts     = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'dd/MM/yyyy HH:mm:ss');
  var rowNum = sheet.getLastRow();

  sheet.appendRow([
    rowNum,
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

  styleStatusCell(sheet, sheet.getLastRow(), 11); // col K = Status

  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + CONFIG.ACCOMMODATION_SPREADSHEET_ID;
  sendWhatsApp('🏠 *NEW ACCOMMODATION — DIT*\n────────────────────\n' +
    '👤 ' + (data.fullName || '') + '\n📧 ' + (data.email || '') + '\n📱 ' + (data.phone || '') +
    '\n🎓 ' + (data.programme || '') + '\n📅 Move-in: ' + (data.moveIn || '') + '\n────────────────────\n📊 ' + sheetUrl);
  sendEmail('🏠 New Accommodation Application: ' + (data.fullName || 'Student') + ' — DIT',
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
    ], sheetUrl));
}

// ── Email builder ───────────────────────────────────────────────

function buildEmailHtml(title, rows, sheetUrl, driveFolderUrl) {
  var rowsHtml = rows.map(function(r) {
    return '<tr><td style="padding:8px 12px 8px 0;color:#555;width:40%;vertical-align:top"><b>' +
           r[0] + '</b></td><td style="padding:8px 0">' + r[1] + '</td></tr>';
  }).join('');
  var driveBtn = driveFolderUrl
    ? ' <a href="' + driveFolderUrl + '" style="background:#1e7e34;color:#fff;padding:10px 20px;border-radius:5px;text-decoration:none;font-size:14px;margin-left:8px;">📁 View Uploaded Files</a>'
    : '';
  return '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;border:1px solid #ddd;border-radius:8px;overflow:hidden">' +
    '<div style="background:#1a2063;padding:20px 24px"><h2 style="color:#fff;margin:0;font-size:18px">' + title + '</h2></div>' +
    '<div style="padding:24px;background:#fff"><table style="width:100%;border-collapse:collapse;font-size:14px">' + rowsHtml + '</table>' +
    '<div style="margin-top:20px"><a href="' + sheetUrl + '" style="background:#1a2063;color:#fff;padding:10px 20px;border-radius:5px;text-decoration:none;font-size:14px">View Spreadsheet</a>' +
    driveBtn + '</div></div></div>';
}

// ── WhatsApp ────────────────────────────────────────────────────

function sendWhatsApp(message) {
  try {
    if (!CONFIG.WA_API_KEY || CONFIG.WA_API_KEY.indexOf('PASTE') !== -1) return;
    UrlFetchApp.fetch('https://api.callmebot.com/whatsapp.php' +
      '?phone='  + encodeURIComponent(CONFIG.WA_NUMBER) +
      '&text='   + encodeURIComponent(message) +
      '&apikey=' + encodeURIComponent(CONFIG.WA_API_KEY));
  } catch (e) { Logger.log('WhatsApp error: ' + e.toString()); }
}

// ── Email ───────────────────────────────────────────────────────

function sendEmail(subject, htmlBody) {
  try {
    MailApp.sendEmail({ to: CONFIG.ADMIN_EMAIL, subject: subject, htmlBody: htmlBody });
  } catch (e) { Logger.log('Email error: ' + e.toString()); }
}

// ── Status cell ─────────────────────────────────────────────────

function styleStatusCell(sheet, row, col) {
  var cell = sheet.getRange(row, col);
  cell.setBackground('#FFF3CD');
  cell.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending','In Review','Contacted','Done','Rejected'], true).build());
}

// ── Setup functions — run each ONCE from the script editor ──────
// These only update headers and column widths, they do NOT delete existing data.

function setupEnrollmentSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.ENROLLMENT_SPREADSHEET_ID);
  var sh = ss.getSheetByName('Enrollments') || ss.insertSheet('Enrollments');
  var hdr = sh.getRange(1, 1, 1, ENR_HEADERS.length);
  hdr.setValues([ENR_HEADERS]).setFontWeight('bold').setFontColor('#ffffff');
  // Color groups: # = grey | info = navy | Programme = blue | admin = green
  sh.getRange(1,1).setBackground('#546e7a');                          // #
  sh.getRange(1,2,1,8).setBackground('#1a1d2e');                      // Timestamp → Nationality
  sh.getRange(1,10).setBackground('#1565c0');                         // Programme (highlight)
  sh.getRange(1,11,1,6).setBackground('#1a1d2e');                     // Edu Level → Reg Form
  sh.getRange(1,17,1,3).setBackground('#2e7d32');                     // Status → Date Checked
  sh.setFrozenRows(1);
  // Column widths: #, Timestamp, Name, Email, Phone, IDType, IDNum, DOB, Nationality,
  //   Programme, EduLevel, HeardAbout, Message, StudentType, HealthDec, RegForm,
  //   Status, AdminNotes, DateChecked
  var w = [45,150,180,210,120,80,140,110,110, 280,140,160,270,100,220,220, 110,240,120];
  w.forEach(function(width, i) { sh.setColumnWidth(i+1, width); });
  Logger.log('Enrollment sheet headers and widths updated!');
}

function setupEnquirySheet() {
  var ss = SpreadsheetApp.openById(CONFIG.ENQUIRY_SPREADSHEET_ID);
  var sh = ss.getSheetByName('Enquiries') || ss.insertSheet('Enquiries');
  var hdr = sh.getRange(1, 1, 1, ENQ_HEADERS.length);
  hdr.setValues([ENQ_HEADERS]).setFontWeight('bold').setFontColor('#ffffff');
  // Color groups: # = grey | info = navy | Message = blue | admin = green
  sh.getRange(1,1).setBackground('#546e7a');                          // #
  sh.getRange(1,2,1,4).setBackground('#1a1d2e');                      // Timestamp → Phone
  sh.getRange(1,6).setBackground('#1565c0');                          // Message (highlight)
  sh.getRange(1,7,1,3).setBackground('#2e7d32');                      // Status → Date Checked
  sh.setFrozenRows(1);
  // Column widths: #, Timestamp, Name, Email, Phone, Message, Status, AdminNotes, DateChecked
  var w = [45,150,180,210,120, 370, 110,240,120];
  w.forEach(function(width, i) { sh.setColumnWidth(i+1, width); });
  Logger.log('Enquiry sheet headers and widths updated!');
}

function setupAccommodationSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.ACCOMMODATION_SPREADSHEET_ID);
  var sh = ss.getSheetByName('Accommodation') || ss.insertSheet('Accommodation');
  var hdr = sh.getRange(1, 1, 1, ACM_HEADERS.length);
  hdr.setValues([ACM_HEADERS]).setFontWeight('bold').setFontColor('#ffffff');
  // Color groups: # = grey | info = navy | Programme = blue | admin = green
  sh.getRange(1,1).setBackground('#546e7a');                          // #
  sh.getRange(1,2,1,7).setBackground('#1a1d2e');                      // Timestamp → ID Number
  sh.getRange(1,9).setBackground('#1565c0');                          // Programme (highlight)
  sh.getRange(1,10).setBackground('#1a1d2e');                         // Move-in Date
  sh.getRange(1,11,1,3).setBackground('#2e7d32');                     // Status → Date Checked
  sh.setFrozenRows(1);
  // Column widths: #, Timestamp, Name, Email, Phone, Gender, StudentType, IDNum,
  //   Programme, MoveIn, Status, AdminNotes, DateChecked
  var w = [45,150,180,210,120,80,110,140, 280,120, 110,240,120];
  w.forEach(function(width, i) { sh.setColumnWidth(i+1, width); });
  Logger.log('Accommodation sheet headers and widths updated!');
}

// ── Accommodation diagnostics — run from script editor ──────────

function testDriveAccess() {
  try {
    Logger.log('Drive folder ID in script: ' + CONFIG.DRIVE_FOLDER_ID);
    if (CONFIG.DRIVE_FOLDER_ID.indexOf('PASTE') !== -1) {
      Logger.log('ERROR: Drive folder ID is still the placeholder. Update CONFIG.DRIVE_FOLDER_ID first.');
      return;
    }
    var folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    Logger.log('Folder name: ' + folder.getName());
    var testFile = folder.createFile('DSH_test_access.txt', 'Drive access test — delete me.', 'text/plain');
    Logger.log('SUCCESS — test file created: ' + testFile.getUrl());
    Logger.log('If you see the file in Drive, the folder is accessible. Delete it manually.');
  } catch (err) {
    Logger.log('ERROR: ' + err.toString());
    Logger.log('This means the script cannot access the Drive folder. Make sure the folder is shared with admissions@dit.edu.my as Editor.');
  }
}

function testAccommodationWrite() {
  try {
    Logger.log('Opening accommodation spreadsheet...');
    var ss = SpreadsheetApp.openById(CONFIG.ACCOMMODATION_SPREADSHEET_ID);
    Logger.log('Spreadsheet name: ' + ss.getName());

    var sheets = ss.getSheets().map(function(s) { return s.getName(); });
    Logger.log('Tabs found: ' + sheets.join(', '));

    var sheet = getOrCreateSheet(ss, 'Accommodation', ACM_HEADERS);
    Logger.log('Sheet found/created: ' + sheet.getName());
    Logger.log('Last row before write: ' + sheet.getLastRow());

    var ts = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'dd/MM/yyyy HH:mm:ss');
    sheet.appendRow([
      sheet.getLastRow(), ts,
      'TEST ENTRY', 'test@test.com', '0100000000',
      'Male', 'Local', 'TEST-000', 'Test Programme', '01/01/2026',
      'Pending', 'Auto-test — delete this row', ''
    ]);
    Logger.log('SUCCESS — row written. Last row now: ' + sheet.getLastRow());
    Logger.log('If you see a test row in the sheet, the script has full access. Delete it manually.');
  } catch (err) {
    Logger.log('ERROR: ' + err.toString());
    Logger.log('This error tells us why the accommodation form is not writing.');
  }
}

function removeAccommodationProtection() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.ACCOMMODATION_SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Accommodation');
    if (!sheet) { Logger.log('No tab named "Accommodation" found.'); return; }
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections = protections.concat(sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE));
    if (protections.length === 0) {
      Logger.log('No protections found on this sheet — protection is not the issue.');
    } else {
      protections.forEach(function(p) { p.remove(); });
      Logger.log('Removed ' + protections.length + ' protection(s). Try submitting the form again.');
    }
  } catch (err) {
    Logger.log('ERROR: ' + err.toString());
  }
}
