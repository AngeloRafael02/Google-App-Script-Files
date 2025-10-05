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