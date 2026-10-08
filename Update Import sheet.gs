function importListFromProductList2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName('Import List');
  const productSheet = ss.getSheetByName('Product List');

  if (!importSheet || !productSheet) {
    SpreadsheetApp.getUi().alert("Please make sure both 'Import List' and 'Product List' sheets exist.");
    return;
  }

  // Force barcode columns to Plain text so leading zeros survive
  importSheet.getRange(4, 1, Math.max(importSheet.getMaxRows() - 3, 1), 1)
             .setNumberFormat('@');
  productSheet.getRange(2, 2, Math.max(productSheet.getMaxRows() - 1, 1), 1)
             .setNumberFormat('@');

  // Get data from Import List starting from row 4
  const importDataRange = importSheet.getRange(4, 1, importSheet.getLastRow() - 3, 7);
  const importData = importDataRange.getValues(); // Columns A to G

  // Barcode as displayed, so a numeric cell doesn't hand back 3764 for "003764"
  const importBarDisp = importSheet
    .getRange(4, 1, importSheet.getLastRow() - 3, 1)
    .getDisplayValues();
  for (let i = 0; i < importData.length; i++) {
    importData[i][0] = importBarDisp[i][0];
  }

  
  // Get data from Product List starting from row 2
  const productData = productSheet.getRange(2, 1, productSheet.getLastRow() - 1, productSheet.getLastColumn()).getValues();

  const prodBarDisp = productSheet
    .getRange(2, 2, productSheet.getLastRow() - 1, 1)
    .getDisplayValues();
  for (let i = 0; i < productData.length; i++) {
    productData[i][1] = prodBarDisp[i][0];
  }

   // Retrieve the data validation rule from 'Product List' for 'Premium Rating' column
  const productPremiumRatingIndex = 9; // Index for 'Premium Rating' in the Product List, Column I
  const importPremiumRatingIndex = 7;  // Index for 'Premium Rating' in the Import List, Column G
  const validationCell = productSheet.getRange(2, productPremiumRatingIndex); // Using row 2 for example
  const validationRule = validationCell.getDataValidation();


  // Define column mappings for Product List
  const productColumns = {
    barcode: 1,        // Column B
    name: 2,           // Column C
    brand: 3,          // Column D
    npp: 4,            // Column E
    mrp: 5,            // Column F
    shellLife: 6,      // Column G
    premiumRating: 8,  // Column I
  };

  // Create a map of Product List using both Barcode and Name as keys
  const productMap = new Map();
  for (let i = 0; i < productData.length; i++) {
    const barcode = productData[i][productColumns.barcode];
    const name = productData[i][productColumns.name];
    const bKey = normBar_(barcode);
    const nKey = String(name == null ? "" : name).trim().toUpperCase();
    if (bKey && !productMap.has(bKey)) productMap.set(bKey, productData[i]);
    if (nKey && !productMap.has(nKey)) productMap.set(nKey, productData[i]);
  }

  // Step 1: Update rows in Import List
  const updatedRows = [];
  const filteredData = []; // Collect rows for re-writing

  for (let i = 0; i < importData.length; i++) {
    const importRow = importData[i];
    if (importRow.every(cell => cell === "")) continue; // Skip blank rows

    const barcode = importRow[0]; // Column A: Item Barcode
    const name = importRow[1];    // Column B: Product Name

    const matchingProduct = productMap.get(normBar_(barcode))
      || productMap.get(String(name == null ? "" : name).trim().toUpperCase());

    if (matchingProduct) {
      importRow[0] = matchingProduct[productColumns.barcode]; // Item Barcode
      importRow[1] = matchingProduct[productColumns.name];    // Product Name
      importRow[2] = matchingProduct[productColumns.brand];   // Brand
      importRow[3] = matchingProduct[productColumns.npp];     // NPP
      importRow[4] = matchingProduct[productColumns.mrp];     // MRP
      importRow[5] = matchingProduct[productColumns.shellLife]; // Shell Life
      importRow[6] = matchingProduct[productColumns.premiumRating]; // Premium Rating
      updatedRows.push(i + 4); // Store the actual row number
    }

    filteredData.push(importRow); // Add row to filtered data
  }

// Apply the data validation to the 'Premium Rating' column in 'Import List'
    if (validationRule) {
      const importValidationRange = importSheet.getRange(4, importPremiumRatingIndex, filteredData.length);
      importValidationRange.setDataValidation(validationRule);
    }

  // Step 2: Overwrite the Import List with the updated data
  importSheet.getRange(4, 1, filteredData.length, 7).setValues(filteredData);
  // Clear rows below the filtered data to remove leftover rows
  const lastRow = importSheet.getLastRow();
  const rowsCleared = lastRow - (filteredData.length + 3); // Count of rows cleared below filtered data
  if (rowsCleared > 0) {
    importSheet.getRange(filteredData.length + 4, 1, rowsCleared, 7).clearContent();
  }

  SpreadsheetApp.getUi().alert(
    `Import complete! Updated rows: ${updatedRows.length}`
  );
}
