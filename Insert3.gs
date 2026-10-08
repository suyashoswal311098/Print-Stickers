/***** GLOBAL CONFIGURATION *****/
const INVENTORY_SS_ID = '1h-Rm2RqJsG_k826d3flemOrYN73laSb5s7ngQn6-E3M'; 

function insert3(e, isSilent = false, cachedProdList = null) {
  // --- Check Conversion Sheet when manually clicking Insert ---
  if (!isSilent) {
    try {
      var ui = SpreadsheetApp.getUi();
      var invSS = SpreadsheetApp.openById(INVENTORY_SS_ID);
      var convSheet = invSS.getSheetByName("Conversion Sheet");
      
      //if (convSheet && convSheet.getLastRow() >= 2) {    //__________________no. of rows in conversion sheet is always greater than 2, even if the cells are blank - sunny
        if (convSheet.getRange("A2").getValue() != "") {   //_________________ checking the value of A2 cell, to determine atleast one entry being available in conversion sheet - sunny
        var response = ui.alert(
          "Pending Inventory Detected", 
          "There are items waiting in the remote Conversion Sheet.\n\nWould you like to pull them now instead?", 
          ui.ButtonSet.YES_NO
        );
        if (response === ui.Button.YES) {
          pullFromInventory(); 
          return "success"; 
        }
      }
    } catch (err) { /* Ignore error and continue */ }
  }

  // 1. Get the Lock
  const lock = LockService.getScriptLock();
  
  // 🚀 THE FIX: Wrapping everything in try...finally guarantees the lock releases
  try {
    const lockSuccess = lock.tryLock(5000); // 5 seconds is safer for bulk processing
    if (!lockSuccess) {
      if (!isSilent) SpreadsheetApp.getUi().alert("Another operation is in progress. Please wait.");
      return "error";
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var work = ss.getSheetByName("Print Console");
    var list = ss.getSheetByName("Product List");

    // Read A2:H2 in one go
    var topRowData = work.getRange("A2:H2").getValues()[0];
    var barcode = topRowData[0]; // A2
    var valueB2 = topRowData[1]; // B2
    var valueC2 = topRowData[2]; // C2
    var valueD2 = topRowData[3]; // D2
    var valueE2 = topRowData[4]; // E2
    var valueF2 = topRowData[5]; // F2
    var valueG2 = topRowData[6]; // G2
    var latestH2 = topRowData[7]; // H2

    var valueC5 = work.getRange("C5").getValue(); 
    var row = Math.max(work.getLastRow() + 1, 8);

    // 🚀 NEW SKIP LOGIC: If Product is not found in List (A2 is #N/A or blank)
    if (!barcode || barcode === "#N/A" || barcode === "#ERROR!") {
      console.warn("Skipping item: '" + valueB2 + "' - Product not found in list.");
      // Clean B2:G2 so the next loop item can paste cleanly
      work.getRange("B2:G2").setValues([["", "", 1, valueE2, "", ""]]);
      return "skipped"; 
    }

    // Standard Validation
    if (!valueB2 || !valueC2) {
      if (!isSilent) SpreadsheetApp.getUi().alert('Kindly fill in Item Code and Sticker Qty');
      return "error";
    }
    if ((valueC5 + valueC2) > 240) {
      if (!isSilent) SpreadsheetApp.getUi().alert('No of stickers exceeding printing limit of 240');
      return "error";
    }

    // Use cached list if provided
        var prodlistData = cachedProdList
      || list.getRange(2, 2, Math.max(list.getLastRow() - 1, 1), 9).getValues();

    for (var i = 0; i < prodlistData.length; i++) {
      var productBarcode = prodlistData[i][0]; 

      if (productBarcode == barcode) {
        
        // Update Product list
        if (valueF2) list.getRange(i + 2, 7).setValue(valueF2); 
        if (valueG2) list.getRange(i + 2, 9).setValue(valueG2); 

        // Build the entire row in memory
        var pkd = "";
        var exp = "";
        if (valueE2 && valueF2) {
          pkd = new Date(valueE2); 
          exp = new Date(pkd.getTime() + (valueF2 * 86400000)); 
        }

        var rowValues = [[productBarcode, "", valueC2, valueD2, pkd, exp, "", "", "", latestH2]];
        var rowColors = [["#ffffff", "#ffffff", "#ffffff", "#ffffff", "#ffffff", "#ffffff", "#ffffff", "#ffffff", "#ffffff", latestH2]];

        // Paste all 10 columns at once
        var targetRange = work.getRange(row, 1, 1, 10);
        targetRange.setValues(rowValues);
        targetRange.setBackgrounds(rowColors);
        
        if (pkd !== "") {
          work.getRange(row, 5, 1, 2).setNumberFormat('DD/MM/YYYY');
        }
        break; 
      }
    }

    // Clean up B2:G2 in ONE single command!
    work.getRange("B2:G2").setValues([["", "", 1, valueE2, "", ""]]);

    // Activate B2 if human is clicking
    if (!isSilent) {
      const fromCheckbox = e && e.range && e.range.getSheet().getName() === 'Print Console' && e.range.getA1Notation() === 'K2';
      if (!fromCheckbox) work.getRange("B2").activate(); 
    }
    
    return "success"; 

  } finally {
    // 🚀 THE FIX: Lock is ALWAYS released, preventing the freeze!
    lock.releaseLock();
  }
}

function pullFromInventory() { 
  const ui = SpreadsheetApp.getUi();
  const startTime = Date.now();
  
  // Connect to Inventory Spreadsheet
  const INVENTORY_SS_ID = '1h-Rm2RqJsG_k826d3flemOrYN73laSb5s7ngQn6-E3M';
  const invSS = SpreadsheetApp.openById(INVENTORY_SS_ID);
  const convSheet = invSS.getSheetByName("Conversion Sheet");
  
  if (!convSheet) return ui.alert("❌ Error: Could not find 'Conversion Sheet'.");
  
  const lastRow = convSheet.getLastRow();
  if (lastRow < 2) return ui.alert("⚠️ Conversion Sheet is empty.");
  
  const data = convSheet.getRange(2, 2, lastRow - 1, 2).getValues(); 
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const work = ss.getSheetByName("Print Console");
  const list = ss.getSheetByName("Product List");
  const cachedProdList = list.getRange(2, 2, Math.max(list.getLastRow() - 1, 1), 9).getValues();
  
  // 🛡️ THE NEW FIX: Read the exact allowed items directly from the B2 dropdown
  const b2Range = work.getRange("B2");
  const b2Validation = b2Range.getDataValidation();
  let allowedValues = [];
  
  if (b2Validation) {
    const criteriaType = b2Validation.getCriteriaType();
    if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      // Extract the allowed names from the dropdown range
      allowedValues = b2Validation.getCriteriaValues()[0].getValues().flat().map(v => String(v).trim());
    } else if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      // Extract allowed names if they were typed manually into the rule
      allowedValues = b2Validation.getCriteriaValues()[0].map(v => String(v).trim());
    }
  }
  
  let itemsProcessed = 0;
  let itemsSkipped = 0;
  let rowsToDelete = [];
  const skipLog = [];
  
  // Loop through Conversion Sheet
  for (let i = 0; i < data.length; i++) {
    // Get the name and trim any invisible spaces
    const productName = String(data[i][0]).trim(); 
    const pkts = data[i][1];    
    const currentRow = i + 2; 
    
    if (!productName) { continue; }

    if (pkts === "" || pkts === null || pkts === undefined || Number(pkts) === 0) {
      skipLog.push(currentRow + ": '" + productName + "' — qty is blank or zero");
      itemsSkipped++;
      convSheet.getRange(currentRow, 2, 1, 2).setBackground("#ffcccc");
      continue;
    }

    // 🛑 PRE-CHECK: If we have a dropdown list, and this product isn't on it, SKIP IT!
    if (allowedValues.length > 0 && !allowedValues.includes(productName)) {
      // Try a forgiving match: collapse whitespace, kill non-breaking spaces,
      // ignore case. Catches the usual paste damage.
      const norm = s => String(s).replace(/\u00A0/g, " ")
                                 .replace(/\s+/g, " ")
                                 .trim()
                                 .toLowerCase();
      const target = norm(productName);
      const loose  = allowedValues.filter(v => norm(v) === target
                                            || norm(v).indexOf(target) >= 0);

      if (loose.length === 1) {
        // Close enough — use the dropdown's own spelling so validation passes.
        data[i][0] = loose[0];
      } else {
        skipLog.push(currentRow + ": '" + productName + "' — not in B2 dropdown"
                     + (loose.length > 1 ? " (" + loose.length + " near matches)" : ""));
        itemsSkipped++;
        convSheet.getRange(currentRow, 2, 1, 2).setBackground("#ffcccc");
        continue;
      }
    }
    
    // ✅ It is on the list! Safe to paste.
    work.getRange("B2:C2").setValues([[data[i][0], pkts]]);
    
    // Wait for VLOOKUPs to load
    SpreadsheetApp.flush();
    Utilities.sleep(500); 
    
    // Call insert3 and check the result status
    const status = insert3(null, true, cachedProdList); 
    
    if (status === "success") {
      itemsProcessed++;
      rowsToDelete.push(currentRow); // Mark for deletion if successful
    } else {
      // If insert3 skipped it for another reason
      itemsSkipped++;
      convSheet.getRange(currentRow, 2, 1, 2).setBackground("#ffcccc");
    }
  }
  
  if (rowsToDelete.length) {
    const ok = ui.alert(
      "Remove imported items?",
      itemsProcessed + " item(s) were inserted successfully.\n\n" +
      "Delete those rows from the Conversion Sheet now?\n" +
      (itemsSkipped ? itemsSkipped + " skipped item(s) will stay, marked red." : ""),
      ui.ButtonSet.YES_NO
    );
    if (ok !== ui.Button.YES) {
      ss.toast("Rows kept. Nothing was deleted from the Conversion Sheet.", "Import done", 6);
      rowsToDelete = [];
    }
  }

  // Clear the successfully-imported rows from the Conversion Sheet.
  // clearContent, NOT deleteRow: A1 holds an ArrayFormula governing this sheet,
  // and deleting rows underneath it can break or shift the spill.
  // Bottom-to-top so nothing shifts mid-loop.
  if (rowsToDelete.length) {
    for (let j = rowsToDelete.length - 1; j >= 0; j--) {
      convSheet.getRange(rowsToDelete[j], 1, 1, 7)
               .clearContent()
               .setBackground(null);
    }
    SpreadsheetApp.flush();
  }

  const timeTaken = ((Date.now() - startTime) / 1000).toFixed(1);
  
  if (skipLog.length) {
    ui.alert(
      "Import finished with skips",
      "Inserted: " + itemsProcessed + "\nSkipped: " + itemsSkipped + "\n\n"
      + skipLog.slice(0, 30).join("\n")
      + (skipLog.length > 30 ? "\n…and " + (skipLog.length - 30) + " more" : ""),
      ui.ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Inserted: ${itemsProcessed} | Skipped: ${itemsSkipped}`,
      "✅ Import Complete",
      8
    );
  }
}