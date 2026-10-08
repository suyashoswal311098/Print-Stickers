function ClearImportProductList() {
  var sheet = SpreadsheetApp.getActiveSheet();

  // 1. Clear rows 3–4 (columns A–G): content, formatting, validations, notes
  sheet.getRange('A3:G4').clear();

  // 2. Delete all rows from row 5 through the very end of the sheet
  var maxRows = sheet.getMaxRows();
  if (maxRows > 4) {
    sheet.deleteRows(5, maxRows - 4);
  }

  // 3. Remove any active filter on the sheet
  var filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }

  // 4. Put the cursor back in A3 ready for the next import
  sheet.getRange('A3').activate();
}
