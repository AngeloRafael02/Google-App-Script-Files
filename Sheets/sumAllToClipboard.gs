/**
 * Calculates the sum of all numerical values within the currently selected cell range(s).
 * Non-numerical cells are ignored. The result is shown in a dialog box for easy copying.
 */
function sumHighlightedCells() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    const range = spreadsheet.getActiveRange();
    if (!range) {
      ui.alert("No cells are currently selected. Please select a range and try again.");
      return;
    }
    const allValues = range.getValues();
    let totalSum = 0;
    
    allValues.forEach(row => {
      row.forEach(cellValue => {
        const number = parseFloat(cellValue);
        if (isFinite(number)) {
          totalSum += number;
        }
      });
    });
    const formattedSum = totalSum.toFixed(2);
    
    ui.alert(`✅ Sum of Highlighted Cells\n\n${formattedSum}`);
  } catch (e) {
    ui.alert(`An error occurred: ${e.message}`);
  }
}