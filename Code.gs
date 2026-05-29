// Paste this entire file into Google Apps Script (Extensions > Apps Script)
// Then deploy as a Web App — see the setup steps in SETUP.md

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Write header row on first submission
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Name', 'Guests', 'Attending']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date().toLocaleString('en-CA', { timeZone: 'America/Toronto' }),
      data.name    || '',
      data.guests  || '',
      data.attending || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Optional: run this once to set up column widths nicely
function setupSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.setColumnWidth(1, 180); // Timestamp
  sheet.setColumnWidth(2, 200); // Name
  sheet.setColumnWidth(3, 100); // Guests
  sheet.setColumnWidth(4, 140); // Attending
  sheet.setName('RSVPs');
}
