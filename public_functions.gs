/**
 * Gets the last row with content in a specific sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to find the last row in.
 * @return {number} - The index of the last row with content.
 */
function getLastRow(sheet) {
  var lastRow = sheet.getLastRow();
  while (lastRow > 0 && sheet.getRange(lastRow, 1).isBlank()) {
    lastRow--;
  }
  return lastRow;
}

/**
 * Sends an email using Gmail service.
 *
 * @param {string} to - The email address of the recipient.
 * @param {string} subject - The subject of the email.
 * @param {string} body - The body of the email.
 */
function sendEmail(to, subject, body) {
  GmailApp.sendEmail(to, subject, body);
}

/**
 * Formats a date object to a string with a specified format.
 *
 * @param {Date} date - The date object to format.
 * @param {string} format - The desired date format (e.g., 'yyyy-MM-dd HH:mm:ss').
 * @return {string} - The formatted date string.
 */
function formatDate(date, format) {
  var formatter = Utilities.formatDate(date, Session.getScriptTimeZone(), format);
  return formatter;
}


/**
 * Gets unique values from a specific column in a sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet containing the column.
 * @param {number} columnIndex - The index of the column to extract unique values from.
 * @return {Array} - An array of unique values.
 */
function getUniqueValues(sheet, columnIndex) {
  var values = sheet.getRange(2, columnIndex, sheet.getLastRow(), 1).getValues();
  var uniqueValues = [...new Set(values.map(function(row) { return row[0]; }))];
  return uniqueValues;
}


/**
 * Gets a sheet by name from a spreadsheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet containing the sheet.
 * @param {string} sheetName - The name of the sheet to retrieve.
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null} - The sheet or null if not found.
 */
function getSheetByName(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  return sheet;
}

/**
 * Merges cells in a specified range in a sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet containing the range.
 * @param {number} startRow - The starting row of the range.
 * @param {number} startColumn - The starting column of the range.
 * @param {number} numRows - The number of rows in the range.
 * @param {number} numColumns - The number of columns in the range.
 */
function mergeCells(sheet, startRow, startColumn, numRows, numColumns) {
  sheet.getRange(startRow, startColumn, numRows, numColumns).merge();
}



/**
 * Extracts hyperlinks from a specified range in a sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet containing the range.
 * @param {number} startRow - The starting row of the range.
 * @param {number} startColumn - The starting column of the range.
 * @param {number} numRows - The number of rows in the range.
 * @param {number} numColumns - The number of columns in the range.
 * @return {Array} - An array of hyperlinks.
 */
function extractHyperlinks(sheet, startRow, startColumn, numRows, numColumns) {
  var range = sheet.getRange(startRow, startColumn, numRows, numColumns);
  var formulas = range.getFormulas();
  var hyperlinks = formulas.map(function(row) {
    return row.map(function(cell) {
      return cell ? cell.match(/=HYPERLINK\("([^"]+)"/) : null;
    });
  });
  return hyperlinks;
}


/**
 * Creates a pivot table in a specified sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to create the pivot table in.
 * @param {string} sourceDataRange - The range of the source data.
 * @param {string} targetCell - The cell where the top-left corner of the pivot table will be placed.
 */
function createPivotTable(sheet, sourceDataRange, targetCell) {
  var pivotTableRange = sheet.getRange(targetCell);
  sheet.insertSheet('Pivot Table');
  var pivotTableSheet = sheet.getSheetByName('Pivot Table');
  var pivotTable = pivotTableSheet.newPivotTable()
    .setValuesSource(sheet.getRange(sourceDataRange))
    .setPosition(pivotTableRange.getRow(), pivotTableRange.getColumn())
    .build();
}



/**
 * Finds and replaces a specified value in a sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to perform find and replace.
 * @param {string} findText - The text to find.
 * @param {string} replaceText - The text to replace found occurrences.
 */
function findAndReplace(sheet, findText, replaceText) {
  var range = sheet.getDataRange();
  var values = range.getValues();
  var newValues = values.map(function(row) {
    return row.map(function(cell) {
      return cell.toString().replace(new RegExp(findText, 'g'), replaceText);
    });
  });
  range.setValues(newValues);
}



/**
 * Protects a specified range in a sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet containing the range.
 * @param {number} startRow - The starting row of the range.
 * @param {number} startColumn - The starting column of the range.
 * @param {number} numRows - The number of rows in the range.
 * @param {number} numColumns - The number of columns in the range.
 */
function protectRange(sheet, startRow, startColumn, numRows, numColumns) {
  var range = sheet.getRange(startRow, startColumn, numRows, numColumns);
  var protection = range.protect().setDescription('Protected Range');
  protection.removeEditors(protection.getEditors());
  protection.addEditor(Session.getActiveUser());
  protection.setWarningOnly(true);

function importExcelData() {
  var excelFile = DriveApp.getFilesByName("Your Excel File Name.xlsx").next();
  var fileId = excelFile.getId();
  var blob = DriveApp.getFileById(fileId).getBlob();
  var data = Utilities.parseCsv(blob.getDataAsString());
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}

}


function syncCalendar() {
  // Replace with your actual calendar ID
  const calendarId = "youremail@gmail.com";
  
  // Replace with the sheet name and data range
  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getRange(4, 1, sheet.getLastRow() - 1, 5); // Skip header row (row 1)
  //const dataValues = dataRange.getValues();
  
  //var cal = CalendarApp.getCalendarById("trudsdata@gmail.com");
  //var events = cal.getEvents(new Date(sht.getRange("B4").getValues()) , new Date(sht.getRange("B5").getValues()));


  // Get Calendar events
  var calendar = CalendarApp.getCalendarById(calendarId);
  var events = calendar.getEvents(new Date(sheet.getRange("G4").getValues()) , new Date(sheet.getRange("G5").getValues()));
  
  // Clear existing sheet data (optional)
  dataRange.clearContent();
  
  // Sync Calendar events to Sheet
  for (let i = 0; i < events.length; i++) {
    const event = events[i];
    dataRange.getCell(i + 1, 5).setValue(event.getTitle()); // Adjust column index based on your data
    dataRange.getCell(i + 1, 4).setValue(event.getDescription()); // Adjust column index based on your data
    dataRange.getCell(i + 1, 2).setValue(event.getEndTime().toString()); // Adjust column index based on your data
    dataRange.getCell(i + 1, 1).setValue(event.getStartTime().toString()); // Adjust column index based on your data

    //dataRange.getCell(i + 1, 3).setValue(event.getEnd().toString()); // Adjust column index based on your data
    // Add more data columns as needed (Description, etc.)

   // var eventTitle = events[eventCtr].getTitle();
    //var eventDesc = events[eventCtr].getDescription(); 
  } 
}
