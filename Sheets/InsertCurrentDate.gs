/*
 * Use to quickly add the current Date to a Selected Cell
 */
function insertDate() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let cell = sheet.getActiveCell(); 
  cell.setValue(new Date()); 
}