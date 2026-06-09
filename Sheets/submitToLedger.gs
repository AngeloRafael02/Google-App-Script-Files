/**
 * gets the values of a transaction from the psuedo-input form in the 
 * Dashboard sheet and submits it to the Ledger sheet as a new entry.
 */
function submitToLedger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName("Dashboard");
  var ledger = ss.getSheetByName("Ledger");
  var ui = SpreadsheetApp.getUi();
  
  var rawDate = String(dashboard.getRange("C3").getValue()).trim();
  var description = String(dashboard.getRange("C4").getValue()).trim();
  var account = String(dashboard.getRange("C5").getValue()).trim();
  var debit = String(dashboard.getRange("C6").getValue()).trim();
  var credit = String(dashboard.getRange("C7").getValue()).trim();
  
  if (description === "" || account === "") {
    ui.alert("⚠️ Incomplete Input", "Description and Account are required fields.", ui.ButtonSet.OK);
    return;
  }
  
  if (debit === "" && credit === "") {
    ui.alert("⚠️ Incomplete Input", "You must enter either a Debit or a Credit amount.", ui.ButtonSet.OK);
    return;
  }
  
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
    
    ui.alert(parts, ui.ButtonSet.OK);
    if (!finalDate) {
      ui.alert("⚠️ Invalid Date Format", "Please enter date as yyyy-mm-dd, yy-mm-dd, yyyy/mm/dd, yy/mm/dd, or 'today'.", ui.ButtonSet.OK);
      return;
    }
  }
  
  var formattedDateStr = Utilities.formatDate(finalDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  ledger.insertRowAfter(1); 
  var targetRange = ledger.getRange("A2:E2");
  targetRange.setValues([[formattedDateStr, description, account, debit, credit]]);
  targetRange.setHorizontalAlignment("center");
  targetRange.setFontWeight("normal");
  targetRange.setFontStyle("normal");
  targetRange.setFontLine("none");

  dashboard.getRange("C4:C7").clearContent();
  ui.alert("Success", "Entry added to Ledger successfully!", ui.ButtonSet.OK);
}