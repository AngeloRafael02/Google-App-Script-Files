/**
 * Used to Aggregate all expenses by Day, used for Charts
 */
function aggregateLedgerData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ledger");
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange("K2:L" + lastRow).clearContent();
  }
  
  var data = sheet.getRange("A2:E" + lastRow).getValues();
  var dailyTotals = {};
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var dateVal = row[0];
    var accountVal = row[2]; // Column C (Account)
    var debitVal = row[3];   // Column D (Debit)
    
    if (dateVal && accountVal) {
      var accountStr = String(accountVal).toLowerCase();
      // Check for 'expense' or 'fee' keywords in Account column
      if (accountStr.indexOf("expense") !== -1 || accountStr.indexOf("fee") !== -1) {
        // Format date to a consistent string key (YYYY-MM-DD)
        var dateStr = Utilities.formatDate(new Date(dateVal), Session.getScriptTimeZone(), "yyyy-MM-dd");
        var debitAmount = parseFloat(debitVal) || 0;
        dailyTotals[dateStr] = (dailyTotals[dateStr] || 0) + debitAmount;
      }
    }
  }
  
  var output = [];
  for (var dateKey in dailyTotals) {
    output.push([dateKey, dailyTotals[dateKey]]);
  }
  
  output.sort(function(a, b) {
    return new Date(a[0]) - new Date(b[0]);
  });
  
  if (output.length > 0) {
    sheet.getRange(2, 11, output.length, 2).setValues(output);
  }
}