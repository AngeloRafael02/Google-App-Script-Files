/**
 * Replaces all non-empty cell values in a fixed range with the number 0.
 * It will not affect conditional formatting or other cell styles.
 * Used in FINANCE Sheet to clear all data
 */
function clearRangeToZero() {
  const SHEET_NAME = "FINANCE"; 
  const TARGET_RANGE = "C21:AG50"; 
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      SpreadsheetApp.getUi().alert(`Error: Sheet named "${SHEET_NAME}" was not found.`);
      return;
    }
    const range = sheet.getRange(TARGET_RANGE);
    const values = range.getValues();
    const newValues = values.map(row => 
      row.map(cellValue => 0)
    );
    range.setValues(newValues);
    SpreadsheetApp.getUi().alert(`Success! All cells in ${TARGET_RANGE} on sheet "${SHEET_NAME}" have been set to 0.`);
  } catch (e) {
    SpreadsheetApp.getUi().alert(`An unexpected error occurred: ${e.toString()}`);
  }
}


/**
 * More Dynamic Version of clearRangeToZero() 
 * which changes cells to zero on active cells of the active sheet
 */
function clearSelectedRangeToZero() {
  const ui = SpreadsheetApp.getUi();

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const range = spreadsheet.getActiveRange(); 
    
    if (!range) {
      ui.alert("Error: No cells or range are currently selected. Please select the range you wish to clear and run the script again.");
      return;
    }
    
    const SHEET_NAME = range.getSheet().getName();
    const TARGET_RANGE = range.getA1Notation(); // e.g., "A1:C10"

    const response = ui.alert(
      'Confirm Action',
      `Are you sure you want to set all cells in the selected range ${TARGET_RANGE} on sheet "${SHEET_NAME}" to 0?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.NO) {
      ui.alert("Operation cancelled by the user.");
      return;
    }

    const values = range.getValues();
    const newValues = values.map(row => 
      row.map(cellValue => 0)
    );
    
    range.setValues(newValues);
    
    ui.alert(`Success! All cells in the selected range ${TARGET_RANGE} on sheet "${SHEET_NAME}" have been set to 0.`);

  } catch (e) {
    ui.alert(`An unexpected error occurred: ${e.toString()}`);
  }
}
