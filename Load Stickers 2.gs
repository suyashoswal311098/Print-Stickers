/**
 * Product Sticker Automation Script
 *
 * Automates sticker generation for products using Google Sheets with the following features:
 *
 * Key Features:
 * 1. **Validation & Alerts**:
 *    - Validates required sticker count (`C5`) and prompts for additional stickers if needed (`C4`).
 * 2. **Sticker Template Management**:
 *    - Dynamically copies a pre-designed template (`A1:F9`) from "#Ref_Novajet 24L Format" to "Novajet 24L."
 *    - Supports grid layout (3 columns x 8 rows per page) with automatic handling of partial pages.
 * 3. **Dynamic Sticker Content**:
 *    - Generates stickers with:
 *      - Product code, name, MRP, NPP, packing/expiry dates, and custom colors.
 * 4. **Batch Number Formatting**:
 *    - Highlights last batch numbers with special formatting (black background, white text, font size 7).
 *    - All other batch numbers use a default font size of 7.
 * 5. **Product List Sync**:
 *    - Updates "Product List" sheet with the latest MRP and NPP values for each product.
 * 6. **Activity Logging**:
 *    - Tracks all generated stickers in the "Print Log" sheet for auditing and transparency.
 *
 * Optimized Layout:
 * - Uses intelligent row and column offsets for seamless grid-based placement.
 * - Automatically manages page transitions and clears unused sections.
 */


 /* function load2() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var work = ss.getSheetByName("Print Console");
    var list = ss.getSheetByName("Product List");
    var stkr = ss.getSheetByName("Novajet 24L");
    var log = ss.getSheetByName("Print Log");
    var rstk = ss.getSheetByName("#Ref_Novajet 24L Format");

    var rl = log.getLastRow() + 1; // Last row for log entry
    var tstr = work.getRange('C5').getValue(); // Total stickers to print
    var rwt = work.getLastRow(); // Total rows in Print Console
    var ui = SpreadsheetApp.getUi();

    if (!tstr || tstr <= 0) {
      ui.alert("Error", "Total stickers (C5) must be greater than zero.", ui.ButtonSet.OK);
      return;
    }

    // Alert if more stickers are needed
    var stickersNeeded = work.getRange("C4").getValue();
    if (stickersNeeded > 0) {
      var response = ui.alert(
        stickersNeeded + " MORE STICKERS NEEDED",
        "Do you want to add " + stickersNeeded + " more stickers to complete the page?",
        ui.ButtonSet.YES_NO
      );

      if (response == ui.Button.YES) {
        var insertPage = ss.getSheetByName("Insert Page"); // Adjust sheet name as needed
        insertPage.appendRow(["Stickers added: " + stickersNeeded]); // Example of adding details
        ui.alert("Stickers added to the 'Insert Page'. Process terminated.");
        return;
      }
    }

// Confirm loading stickers
var response = ui.alert("LOAD STICKERS", "Do you want to Load the Stickers?", ui.ButtonSet.YES_NO);
if (response == ui.Button.NO) return;

// Calculate the total number of stickers required and track remaining stickers
var stickersRemaining = tstr; // Total stickers to print
var rowsPerSticker = 9; // Number of rows per sticker
var columnsPerSticker = 6; // Number of columns per sticker
var stickersPerPage = 24; // Total stickers per page (3 columns x 8 rows of stickers)

// Load stickers individually
for (var i = 1; stickersRemaining > 0; i++) {
  // Load template for the current sticker
  for (var j = 0; j < Math.min(stickersRemaining, stickersPerPage); j++) {
    var stickerRow = Math.floor(j / 3) * rowsPerSticker + (i - 1) * stickersPerPage * 3; // Row offset for each sticker
    var stickerColumn = (j % 3) * columnsPerSticker + 1; // Column offset for each sticker

    // Copy the template for each sticker
    rstk.getRange('A1:F9').copyTo(
      stkr.getRange(stickerRow + 1, stickerColumn, rowsPerSticker, columnsPerSticker),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );
  }

  // Clear any unused stickers on the last page
  if (stickersRemaining < stickersPerPage) {
    var startClearing = Math.min(stickersRemaining, stickersPerPage); // Stickers already placed
    for (var j = startClearing; j < stickersPerPage; j++) {
      var stickerRow = Math.floor(j / 3) * rowsPerSticker + (i - 1) * stickersPerPage * 3; // Row offset for unused stickers
      var stickerColumn = (j % 3) * columnsPerSticker + 1; // Column offset for unused stickers
      stkr.getRange(stickerRow + 1, stickerColumn, rowsPerSticker, columnsPerSticker).clearContent();
    }
  }

  // Reduce remaining stickers
  stickersRemaining -= stickersPerPage;
}

  
// Prepare for generating stickers
var stickerDetails = {
  stc: 1, // Sticker column (1, 2, or 3)
  ic: 7,  // Start from row 1 instead of row 7
  iname: 4,
  ib: 3,
  imrp: 5,
  inpp: 6,
  idt: 9,
  inote: 8,
  iseperator: 1,
  currentNumber: 1 // Start numbering from 1 for the first product batch
};

// Generate stickers for each product
for (var rw = 8; rw <= rwt; rw++) {
  var rowDetails = getRowDetails(work, rw);
  if (!rowDetails) continue; // Skip invalid rows

  updateProductList(list, rowDetails);

  // Reset productCount for each new product row
  stickerDetails.productCount = 0; // Reset numbering to start at 1 for this product batch

  for (var k = 1; k <= rowDetails.sn; k++) {
  addStickerWithNumber(stkr, stickerDetails, rowDetails, k, rowDetails.sn); // Pass batch number and total batches
  stickerDetails.currentNumber++; // Increment the global sticker number
}


  addLogEntry(log, rl++, rowDetails);
}

  }

  // Get details of the current row
  function getRowDetails(sheet, row) {
    var sn = sheet.getRange(row, 3).getValue(); // Stickers needed
    if (!sn || sn <= 0) return null;

    return {
      sn,
      wcode: sheet.getRange(row, 1).getValue(),
      wname: sheet.getRange(row, 2).getValue(),
      wpack: sheet.getRange(row, 4).getValue(),
      wbrand: sheet.getRange(row, 7).getValue(),
      wnpp: sheet.getRange(row, 8).getValue(),
      wmrp: sheet.getRange(row, 9).getValue(),
      wnote: sheet.getRange(row, 11).getValue(),
      clr: sheet.getRange(row, 10).getValue() || "#FFFFFF",
      pk: new Date(sheet.getRange(row, 5).getValue()),
      ex: new Date(sheet.getRange(row, 6).getValue())
    };
  }

  // Update product list with current details
  function updateProductList(list, details) {
    var productList = list.getRange("A1:J3000").getValues();
    for (var i = 0; i < productList.length; i++) {
      if (productList[i][1] == details.wcode) {
        var j = i + 1;
        list.getRange(j, 5).setValue(details.wnpp); // Update NPP
        list.getRange(j, 6).setValue(details.wmrp); // Update MRP
        break;
      }
    }
  }

function addStickerWithNumber(sheet, stickerDetails, details, batchNumber, totalBatches) {
  var { stc, ic, iname, ib, imrp, inpp, idt, iseperator } = stickerDetails;

  // Calculate the column offset
  var colOffset = (stc - 1) * 6; // Offset for 1st, 2nd, or 3rd column in the grid

  // Construct barcode with wpack if wpack is not 1
  var barcode = details.wpack !== 1 ? `${details.wcode}-${details.wpack}` : details.wcode;

  // Place product code (barcode)
  sheet.getRange(ic, 3 + colOffset).setValue(barcode);

  // Place product name
  sheet.getRange(iname, 1 + colOffset)
    .setValue(details.wname)
    .setFontColor(details.clr === "#FFFFFF" ? "#000000" : "#ffffff")
    .setBackground(details.clr);

    // Adjust MRP and NPP values based on wpack and round to the nearest whole number
    var adjustedMrp = details.wpack !== 1 ? Math.round(details.wmrp * details.wpack) : Math.round(details.wmrp);
    var adjustedNpp = details.wpack !== 1 ? Math.round(details.wnpp * details.wpack) : Math.round(details.wnpp);

    // Construct MRP and NPP display text
    var mrpText = `₹${adjustedMrp}/-`; // Round figure for MRP
    var nppText = `₹${adjustedNpp}/-`; // Round figure for NPP

    // Place MRP
    sheet.getRange(imrp, 2 + colOffset).setValue(mrpText);

    // Place NPP
    sheet.getRange(inpp, 2 + colOffset).setValue(nppText);


  // Place packing and expiry dates
  if (details.pk && details.ex) {
    sheet.getRange(idt, 1 + colOffset).setValue(`PKD: ${details.pk.getDate()}/${details.pk.getMonth() + 1}/${details.pk.getFullYear()}`);
    sheet.getRange(idt, 4 + colOffset).setValue(`EXP: ${details.ex.getDate()}/${details.ex.getMonth() + 1}/${details.ex.getFullYear()}`);
  }

  // Place batch number
  var batchNumberRow = iseperator;
  var batchNumberCol = 6 + colOffset; // Column F
  var batchCell = sheet.getRange(batchNumberRow, batchNumberCol);
  batchCell.setValue(batchNumber);

  // Apply default font size for all batch numbers
  batchCell.setFontSize(7);

  // Apply special formatting for last batch number
  if (batchNumber === totalBatches) {
    batchCell.setFontColor("#FFFFFF") // White text
      .setBackground("#000000") // Black background
      .setFontSize(7); // Font size 7
  }

  // Update sticker column and row offsets
  stickerDetails.stc = stc % 3 + 1; // Cycle through columns
  if (stc === 3) adjustRowOffsets(stickerDetails); // Move to the next block of rows
}


  // Adjust row offsets after three stickers
  function adjustRowOffsets(details) {
    details.ic += 9;
    details.iname += 9;
    details.ib += 9;
    details.imrp += 9;
    details.inpp += 9;
    details.idt += 9;
    details.inote += 9;
    details.iseperator += 9;
  }

  // Add a log entry for printed stickers
  function addLogEntry(log, rl, details) {
    log.getRange(rl, 1).setValue(details.wcode);
    log.getRange(rl, 2).setValue(details.wname);
    log.getRange(rl, 3).setValue(details.sn);
    log.getRange(rl, 4).setValue(details.wpack);
    log.getRange(rl, 5).setValue(`${details.pk.getDate()}/${details.pk.getMonth() + 1}/${details.pk.getFullYear()}`);
    log.getRange(rl, 6).setValue(`${details.ex.getDate()}/${details.ex.getMonth() + 1}/${details.ex.getFullYear()}`);
    log.getRange(rl, 7).setValue(details.wbrand);
    log.getRange(rl, 8).setValue(details.wmrp);
    log.getRange(rl, 9).setValue(details.wnpp);
    log.getRange(rl, 10).setValue(details.clr);
    log.getRange(rl, 11).setValue(new Date()).setNumberFormat('dd-MM-yyyy hh:mm');
    log.getRange(rl, 12).setValue(Session.getActiveUser().getEmail());
  }*/


