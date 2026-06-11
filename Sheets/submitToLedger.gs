/**
 * gets the values of a transaction from the pseudo-input form in the 
 * Dashboard sheet and submits it to the Ledger sheet as a new entry.
 */
function submitToLedger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ledger = ss.getSheetByName("Ledger");
  var ui = SpreadsheetApp.getUi();
  
  var rawDate = String(ledger.getRange("I4").getValue()).trim();
  var description = String(ledger.getRange("I5").getValue()).trim();
  var account = String(ledger.getRange("I6").getValue()).trim(); // For the Credit Row
  var expenseType = String(ledger.getRange("I7").getValue()).trim();  // Expense Type -> For the Debit Row
  var amount = String(ledger.getRange("I8").getValue()).trim();
  
  if (description === "" || account === "" || expenseType === "") {
    ui.alert("⚠️ Incomplete Input", "Description, Account, and Expense Type are required fields.", ui.ButtonSet.OK);
    return;
  }
  
  if (amount === "" || isNaN(parseFloat(amount)) || parseFloat(amount) <= 0) {
    ui.alert("⚠️ Invalid Amount", "Please enter a valid amount greater than 0.", ui.ButtonSet.OK);
    return;
  }
  
  // --- Date Parsing Logic ---
  var finalDate;  
  if (rawDate === "" || rawDate.toLowerCase() === "today") {
    finalDate = new Date();
  } else {
    var cleanDateStr = rawDate.replace(/\//g, "-");
    var parts = cleanDateStr.split("-");
    
    if (parts.length === 3) {
      var year = parts[0];
      var month = parseInt(parts[1], 10) - 1; // JS Months are 0-11
      var day = parseInt(parts[2], 10);

      if (year.length === 2) {
        // Assumes '20xx' for 2-digit years
        year = "20" + year;
      }

      year = parseInt(year, 10);
      var parsedDate = new Date(year, month, day);
      
      if (!isNaN(parsedDate.getTime()) && parsedDate.getFullYear() === year && parsedDate.getMonth() === month && parsedDate.getDate() === day) {
        finalDate = parsedDate;
      }
    }
    
    if (!finalDate) {
      ui.alert("⚠️ Invalid Date Format", "Please enter date as yyyy-mm-dd, yy-mm-dd, yyyy/mm/dd, yy/mm/dd, or 'today'.", ui.ButtonSet.OK);
      return;
    }
  }
  
  var formattedDateStr = Utilities.formatDate(finalDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  // --- Writing Dual Entries to Ledger ---
  
  // Style resetting helper function to keep ledger clean
  function applyRowFormatting(range) {
    range.setHorizontalAlignment("center");
    range.setFontWeight("normal");
    range.setFontStyle("normal");
    range.setFontLine("none");
  }

  ledger.insertRowAfter(1); 
  var creditRange = ledger.getRange("A2:E2");
  creditRange.setValues([[formattedDateStr, description, account, "", amount]]);
  applyRowFormatting(creditRange);

  ledger.insertRowAfter(1); 
  var debitRange = ledger.getRange("A2:E2");
  debitRange.setValues([[formattedDateStr, description, expenseType, amount, ""]]);
  applyRowFormatting(debitRange);

  ledger.getRange("I5:I8").clearContent();
  ui.alert("Success", "Double-entry transaction added successfully!", ui.ButtonSet.OK);
}