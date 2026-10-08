/**
 * Product Sticker Automation Script
 *
 * This script automates the process of generating and managing product stickers within Google Sheets, leveraging data from associated product lists to ensure accurate and up-to-date sticker content.
 *
 * Key Features:
 * 1. **Validation & Alerts**:
 *    - Ensures that the total sticker count specified in 'C5' is positive, and prompts the user to add additional stickers if the count in 'C4' is insufficient.
 *
 * 2. **Sticker Template Management**:
 *    - Copies a pre-defined sticker template from the "#Ref_Novajet 24L Format" sheet to the "Novajet 24L" sheet, accommodating different layouts including partial pages.
 *
 * 3. **Dynamic Sticker Content**:
 *    - Automatically fills stickers with essential product information such as product code, name, Manufacturer's Recommended Price (MRP), Net Purchase Price (NPP), packing and expiry dates, and applies custom color formatting.
 *
 * 4. **Batch Number Formatting**:
 *    - Applies special formatting to the last batch number of each product series to enhance visibility and differentiation. Standard batch numbers maintain a uniform font size throughout.
 *
 * 5. **Product List Sync**:
 *    - Updates the "Product List" sheet dynamically to reflect the most recent MRP and NPP values directly from the sticker generation process.
 *
 * 6. **Activity Logging**:
 *    - Maintains a detailed log of all sticker generation activities in the "Print Log" sheet, providing an audit trail for transparency and record-keeping.
 *
 * Optimized Layout:
 * - Implements a grid-based sticker layout strategy, efficiently managing space to ensure optimal placement and page utilization. Unused sticker spaces are automatically cleared to maintain presentation integrity.
 *
 * The script enhances operational efficiency by automating complex tasks, ensuring data consistency, and facilitating easy monitoring and adjustments as needed.
 */


  function load2() {
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

var stickersNeeded = work.getRange("C4").getValue();
if (stickersNeeded > 0) {
  var response = ui.alert(
    stickersNeeded + " MORE STICKERS NEEDED",
    "Do you want to add " + stickersNeeded + " more stickers to complete the page?",
    ui.ButtonSet.YES_NO
  );

  if (response != ui.Button.NO) { 
      ui.alert("Add these stickers: " + stickersNeeded);
    return;
  }
}

// Second Alert (Load Stickers)
var response = ui.alert("LOAD STICKERS", "Do you want to Load the Stickers?", ui.ButtonSet.YES_NO_CANCEL);
if (response != ui.Button.YES) { // Stops execution if dismissed
  return;
}

// Third Alert (Data Present on Sticker Sheet)
var stickerspresent = stkr.getRange("B1").getValue();
if (stickerspresent != "") {
var response = ui.alert("DATA PRESENT!!!", "Do you want to overwrite the Stickers?", ui.ButtonSet.YES_NO);
if (response != ui.Button.YES)// Stops execution if dismissed
  return;
}

// Continue only if YES is selected.

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
var productIndex = 0; // Counter for products
for (var rw = 8; rw <= rwt; rw++) {
  var rowDetails = getRowDetails(work, rw);
  if (!rowDetails) continue; // Skip invalid rows
  
  // Set labelFirstSticker: for the first product, true; for others, false.
  productIndex++;
  rowDetails.labelFirstSticker = (productIndex === 1);
  
  updateProductList(list, rowDetails);

  // Reset numbering for each new product row
  stickerDetails.productCount = 0; // Reset numbering to start at 1 for this product batch

  for (var k = 1; k <= rowDetails.sn; k++) {
    addStickerWithNumber(stkr, stickerDetails, rowDetails, k, rowDetails.sn); // Pass batch number and total batches
    stickerDetails.currentNumber++; // Increment the global sticker number
  }

  addLogEntry(log, rl++, rowDetails);
}

  }

function getRowDetails(sheet, row) {
    var sn = sheet.getRange(row, 3).getValue(); // Stickers needed
    if (!sn || sn <= 0) return null;

    var wpack = sheet.getRange(row, 4).getValue();
    var wmrp = sheet.getRange(row, 9).getValue();
    var wnpp = sheet.getRange(row, 8).getValue();
    var wcode = sheet.getRange(row, 1).getValue();
    var wname = sheet.getRange(row, 2).getValue();

    // Initialize adjusted values with original ones
    var adjustedMrp = wmrp;
    var adjustedNpp = wnpp;

    // Extract the weight from the name
    var weightMatch = wname.match(/(\d+)(\s*)(gm)/i); // Match the number and "gm" (case-insensitive)
    var weight = weightMatch ? parseFloat(weightMatch[1]) : null;

    // Adjust wmrp, wnpp, and optionally weight if wpack is less than 1 or greater than 1
    if (wpack < 1 || wpack > 1 && weight) {
        adjustedMrp = wmrp * wpack;
        adjustedNpp = wnpp * wpack;

        // Modify wcode to include "-wpack"
        wcode += `-${wpack}`;

        // Automatically calculate new weight
        var newWeight = Math.round(weight * wpack); // Calculate adjusted weight
        wname = wname.replace(weightMatch[0], `${newWeight}GM`); // Replace old weight with new weight and "GM"
    }

    return {
        sn,
        wcode,
        wname, // Updated wname with modified weight and uppercase "GM"
        wpack,
        wbrand: sheet.getRange(row, 7).getValue(),
        wnpp: adjustedNpp,
        wmrp: adjustedMrp,
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

  // Calculate column offset (for 1st, 2nd, or 3rd column in the grid)
  var colOffset = (stc - 1) * 6;

  // Place product code
  sheet.getRange(ic, 3 + colOffset).setValue(details.wcode);

  // Place product name with its background and font colors
  sheet.getRange(iname, 1 + colOffset)
    .setValue(details.wname)
    .setFontColor(details.clr === "#FFFFFF" ? "#000000" : "#ffffff")
    .setBackground(details.clr);

  // Place adjusted MRP (rounded to nearest integer)
  sheet.getRange(imrp, 2 + colOffset).setValue("₹" + Math.round(details.wmrp) + "/-");

  // Place adjusted NPP (rounded to nearest integer)
  sheet.getRange(inpp, 2 + colOffset).setValue("₹" + Math.round(details.wnpp) + "/-");

  // Place brand name
  sheet.getRange(ib, 1 + colOffset).setValue(details.wbrand);

  // Place packing and expiry dates if available
  if (details.pk && details.ex) {
    sheet.getRange(idt, 1 + colOffset)
         .setValue(`PKD: ${details.pk.getDate()}/${details.pk.getMonth() + 1}/${details.pk.getFullYear()}`);
    sheet.getRange(idt, 4 + colOffset)
         .setValue(`EXP: ${details.ex.getDate()}/${details.ex.getMonth() + 1}/${details.ex.getFullYear()}`);
  }

  // Determine the cell for batch number display (Column F)
  var batchNumberRow = iseperator;
  var batchNumberCol = 6 + colOffset;
  var batchCell = sheet.getRange(batchNumberRow, batchNumberCol);

  // Only label the last sticker of the batch
  if (batchNumber === totalBatches) {
    batchCell.setValue(totalBatches)
             .setFontColor("#FFFFFF")
             .setBackground("#000000")
             .setFontSize(7);
  } else {
    batchCell.setValue("");
  }

  // Update sticker column and row offsets (cycle through columns and adjust rows when needed)
  stickerDetails.stc = stc % 3 + 1;
  if (stc === 3) adjustRowOffsets(stickerDetails);
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
  }


