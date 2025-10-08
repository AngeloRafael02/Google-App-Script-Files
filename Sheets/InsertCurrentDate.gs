function insertDate() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let cell = sheet.getActiveCell(); 
  cell.setValue(new Date()); 
}