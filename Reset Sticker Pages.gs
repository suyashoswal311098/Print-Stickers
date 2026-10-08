/**
 * Reset Sticker Pages Script
 *
 * This script resets the "Novajet 24L" sheet by clearing all rows, columns, content, and formatting.
 *
 * Key Features:
 * 1. **Confirmation Prompt**:
 *    - Prompts the user to confirm the reset action to prevent accidental resets.
 * 2. **Full Sheet Clearing**:
 *    - Clears all rows and columns, including empty cells, resetting the sheet entirely.
 */

function reset() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var work = ss.getSheetByName("Novajet 24L");

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("RESET STICKER PAGES", "Do you want to Reset the Sticker Pages?", ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    // Get the maximum rows and columns of the sheet
    var maxRows = work.getMaxRows();
    var maxColumns = work.getMaxColumns();

    // Define the range to clear
    var rangeToClear = work.getRange(1, 1, maxRows, maxColumns);

    // Clear all content and formatting
    rangeToClear.clear(); // Clears all content and resets all formatting
  }
}
