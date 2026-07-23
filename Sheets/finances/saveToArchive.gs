/**
 * Moves monthly Expenses Data to Archives Sheet.
 * Used in Expenses Sheet
 */
function SaveToArchive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expensesSheet = ss.getSheetByName('Expenses');
  const archivesSheet = ss.getSheetByName('Archives');
  
  const currentMonth = expensesSheet.getRange('B2').getValue().toString().trim();
  
  if (!currentMonth) {
    SpreadsheetApp.getUi().alert('No month found in cell B2 of the Expenses sheet.');
    return;
  }
  
  const sourceRange = expensesSheet.getRange('C3:C12');
  const valuesToArchive = sourceRange.getValues();
  
  const archiveMonths = archivesSheet.getRange('C3:N3').getValues()[0]; 
  
  let targetColIndex = -1;
  for (let i = 0; i < archiveMonths.length; i++) {
    if (archiveMonths[i].toString().trim().toUpperCase() === currentMonth.toUpperCase()) {
      targetColIndex = i + 3;
      break;
    }
  }
  
  if (targetColIndex !== -1) {
    archivesSheet.getRange(4, targetColIndex, valuesToArchive.length, 1).setValues(valuesToArchive);
    SpreadsheetApp.getActiveSpreadsheet().toast('Data successfully archived for ' + currentMonth + '!', 'Success');
  } else {
    SpreadsheetApp.getUi().alert('Could not find a column for "' + currentMonth + '" in the Archives sheet.');
  }
}