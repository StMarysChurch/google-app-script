function clearSheet() {
  // Clears the sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var range = sheet.getRange("B5:C14");
  range.clearContent();
  sheet = ss.getSheets()[1];
  range = sheet.getRange("A2:D200");
  range.clearContent();

  sheet = ss.getSheets()[2];
  range = sheet.getRange("A2:D200");
  range.clearContent();

  sheet = ss.getSheets()[3];
  range = sheet.getRange("A2:D200");
  range.clearContent();

  sheet = ss.getSheets()[4];
  range = sheet.getRange("A2:D200");
  range.clearContent();

  sheet = ss.getSheets()[5];
  range = sheet.getRange("A2:D200");
  range.clearContent();

  sheet = ss.getSheets()[6];
  range = sheet.getRange("A2:D200");
  range.clearContent();

  sheet = ss.getSheets()[7];
  range = sheet.getRange("A2:D200");
  range.clearContent();

  sheet = ss.getSheets()[8];
  range = sheet.getRange("A2:D200");
  range.clearContent();
}
