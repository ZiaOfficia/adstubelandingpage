// AdsTube — Google Apps Script Web App
// Steps to deploy:
//   1. Go to your Google Sheet → Extensions → Apps Script
//   2. Paste this entire file, replacing any existing code
//   3. Click Deploy → New deployment → Web app
//      Execute as: Me | Who has access: Anyone
//   4. Copy the Web App URL and paste it into index.html (SCRIPT_URL constant)

var SHEET_NAME = 'Leads';

function doPost(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    // Write header row once
    if (sheet.getLastRow() === 0) {
      var headers = [
        'Timestamp',
        'Form',
        'Triggered By',
        'Name',
        'Phone / WhatsApp',
        'Email',
        'YouTube Channel URL',
        'Service',
        'Budget',
        'Message / Goal'
      ];
      sheet.appendRow(headers);
      var hRange = sheet.getRange(1, 1, 1, headers.length);
      hRange.setFontWeight('bold')
            .setBackground('#e8001c')
            .setFontColor('#ffffff')
            .setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 160);  // Timestamp
      sheet.setColumnWidth(2, 160);  // Form
      sheet.setColumnWidth(3, 200);  // Triggered By
      sheet.setColumnWidth(4, 140);  // Name
      sheet.setColumnWidth(5, 150);  // Phone
      sheet.setColumnWidth(6, 180);  // Email
      sheet.setColumnWidth(7, 240);  // Channel URL
      sheet.setColumnWidth(8, 160);  // Service
      sheet.setColumnWidth(9, 130);  // Budget
      sheet.setColumnWidth(10, 280); // Message
    }

    var p = e.parameter;

    sheet.appendRow([
      new Date(),
      p.form        || 'Unknown Form',
      p.triggeredBy || 'Unknown',
      p.name        || '',
      p.phone       || '',
      p.email       || '',
      p.channelUrl  || '',
      p.service     || '',
      p.budget      || '',
      p.message     || ''
    ]);

    // Auto-resize rows to prevent overflow
    sheet.autoResizeRows(sheet.getLastRow(), 1);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Health-check — visit the Web App URL in browser to confirm it's live
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ result: 'ok', service: 'AdsTube Lead Collector' }))
    .setMimeType(ContentService.MimeType.JSON);
}
