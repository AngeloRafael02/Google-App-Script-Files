/**
 * Clears ACCOUNTING Sheet and refetched data from
 * FINANCE to create a new balance sheet
 */
function updateAccountingSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var financeSheet = ss.getSheetByName("FINANCE");
  var accountingSheet = ss.getSheetByName("ACCOUNTING");

  accountingSheet.getRange("A2:G").clearContent();

  var month = financeSheet.getRange("B19").getValue();
  var days = financeSheet.getRange("C20:AG20").getValues()[0];
  var categories = financeSheet.getRange("A21:A53").getValues();
  var descriptions = financeSheet.getRange("B21:B53").getValues();
  var data = financeSheet.getRange("C21:AG53").getValues();

  var outputData = [];
  var runningBalance = 0;

  for (var col = 0; col < data[0].length; col++) {
    var day = days[col];
    if (!day || day === "TOTAL") continue;

    var currentCategory = "";
    for (var row = 0; row < data.length; row++) {
      if (categories[row][0] !== "") {
        currentCategory = categories[row][0];
      }
      
      var amount = data[row][col];
      if (amount !== "" && amount !== 0 && amount !== null) {
        var dateStr = month + " " + day;
        var description = descriptions[row][0];
        runningBalance += Number(amount) || 0;

        outputData.push([dateStr, description, currentCategory, "", "", amount, runningBalance]);
      }
    }
  }

  if (outputData.length > 0) {
    accountingSheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
  }
}