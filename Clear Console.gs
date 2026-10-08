function Clearconsole() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("CLEAR CONSOLE", "Do you want to Reset the Console Data?", ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var consoleSheet = spreadsheet.getSheetByName("Print Console");
    var formattedToday = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    if (consoleSheet) {
      var lastRow = consoleSheet.getLastRow();

      if (lastRow >= 8) {
        consoleSheet.deleteRows(8, lastRow - 7);
      }

      consoleSheet.getRangeList(['B2','C2','F4','C5', 'C4','E2']).clearContent();
      consoleSheet.getRange('C5').setFormula("=Sum($C$7:$C)");
      consoleSheet.getRange('F4').setFormula("=QUOTIENT($C$5,24)");
      consoleSheet.getRange('C4').setFormula('=if(MOD($C$5,24)=0,"0",(24-MOD($C$5,24)))');
      consoleSheet.getRange('E2').setValue(formattedToday);

      SpreadsheetApp.flush();
    }
  }
}


function setTodaysDateInWorkSheet() {
  SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Print Console")
    .getRange("E2")
    .setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));
}

function createOnOpenTrigger() {
  ScriptApp.newTrigger("setTodaysDateInWorkSheet")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onOpen()
    .create();
}
