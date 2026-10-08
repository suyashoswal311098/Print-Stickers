function executeWithConfirmation(funcName) {
  const ui = SpreadsheetApp.getUi();

  // 1️⃣ Ask for confirmation
  const response = ui.alert(
    "Start Process",
    `Do you want to proceed with the ${funcName === 'importListFromProductList2' 
       ? 'Import Products' 
       : funcName === 'updateProductList2' 
         ? 'Update Products' 
         : 'Unknown'} process? This will modify your data.`,
    ui.ButtonSet.YES_NO
  );

  // 2️⃣ If not YES, bail out
  if (response !== ui.Button.YES) {
    ui.alert("Process canceled.");
    return;
  }

  // 3️⃣ Only YES comes here
  switch (funcName) {
    case "importListFromProductList2":
      importListFromProductList2();
      break;
    case "updateProductList2":
      updateProductList2();
      break;
    default:
      ui.alert("Invalid function name provided.");
  }
}


function executeImport() {
  executeWithConfirmation("importListFromProductList2");
}

function executeUpdate() {
  executeWithConfirmation("updateProductList2");
}

/**
 * EXPORT PRODUCTS TO XLS
 *
 * Asks which sheet to export from, then exports EXACTLY what is visible:
 * rows hidden by a filter or hidden manually are skipped. Clear the filter
 * to get everything.
 *
 * Headers: Code | Product Name | Article Code | Price | MRP | Variants_Name | Variants_Price
 *
 * Barcodes are read with getDisplayValues() and written into a text-formatted
 * cell, so 003764 stays 003764 in the file instead of collapsing to 3764.
 * The portal's Update Price upload rejects .csv and .xlsx, so this writes a real
 * Excel 97-2003 .xls with SheetJS (bundled in SheetJS.js; Apps Script can't write .xls).
 */

// ---- COLUMN MAPPING PER SHEET -------------------------------------------
// Use a column letter, or "" to leave that output column blank.
const EXPORT_MAP_ = {
  'Product List': {
    firstDataRow:  2,
    code:          'B',
    productName:   'C',
    articleCode:   '',    // <-- set if you have one
    price:         'E',   // NPP
    mrp:           'F',
    variantsName:  '',    // <-- set if you have one
    variantsPrice: ''     // <-- set if you have one
  },
  'Import List': {
    firstDataRow:  4,
    code:          'A',
    productName:   'B',
    articleCode:   '',
    price:         'D',   // NPP
    mrp:           'E',
    variantsName:  '',
    variantsPrice: ''
  }
};

// Columns that must keep their leading zeros (stored as text).
const TEXT_FIELDS_ = ['code', 'articleCode'];
// -------------------------------------------------------------------------


function exportProductsToXlsx() {
  const ui = SpreadsheetApp.getUi();

  // --- Ask which sheet ---------------------------------------------------
  const resp = ui.prompt(
    "Export to XLS",
    "Type 1 or 2, then press OK:\n\n" +
    "   1  =  Import List\n" +
    "   2  =  Product List\n\n" +
    "Only the rows currently VISIBLE are exported. If a filter is on you get " +
    "the filtered rows only — clear the filter first if you want everything.",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const pick = String(resp.getResponseText()).trim();
  let sheetName;
  if (pick === '1')      sheetName = 'Import List';
  else if (pick === '2') sheetName = 'Product List';
  else { ui.alert("Please type 1 or 2. Nothing was exported."); return; }

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(sheetName);
  if (!src) { ui.alert("Sheet not found: " + sheetName); return; }

  const map = EXPORT_MAP_[sheetName];
  const firstRow = map.firstDataRow;
  const nRows = src.getLastRow() - firstRow + 1;
  if (nRows <= 0) { ui.alert("No data on " + sheetName + "."); return; }

  const lastCol  = src.getLastColumn();
  const values   = src.getRange(firstRow, 1, nRows, lastCol).getValues();
  const displays = src.getRange(firstRow, 1, nRows, lastCol).getDisplayValues();

  // --- Work out which rows are visible ----------------------------------
  const visible = visibleRowFlags_(src, firstRow, nRows);

  const order   = ['code','productName','articleCode','price','mrp','variantsName','variantsPrice'];
  const headers = ['Code','Product Name','Article Code','Price','MRP','Variants_Name','Variants_Price'];

  let skipped = 0;
  const out = [headers];
  for (let i = 0; i < values.length; i++) {
    if (!visible[i]) { skipped++; continue; }

    const row = order.map(field => {
      const letter = map[field];
      if (!letter) return "";
      const idx = colLetterToIndex_(letter);
      if (idx < 0 || idx >= lastCol) return "";
      return TEXT_FIELDS_.indexOf(field) >= 0
        ? String(displays[i][idx]).trim()   // display value keeps the zeros
        : values[i][idx];
    });
    if (row.some(v => v !== "" && v !== null)) out.push(row);
  }

  if (out.length < 2) { ui.alert("No visible rows to export."); return; }

  // --- Build a real .xls (Excel 97-2003) with the bundled SheetJS --------
  const stamp   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
  const tmpName = sheetName.replace(/\s+/g, "") + "_Export_" + stamp;

  // Codes are strings, so leading zeros survive. Blanks become empty cells.
  const aoa = out.map(r => r.map(v => (v === "" || v === undefined) ? null : v));
  const lib = sheetJs_();
  const wb  = lib.utils.book_new();
  lib.utils.book_append_sheet(wb, lib.utils.aoa_to_sheet(aoa), "Worksheet");
  const u8    = new Uint8Array(lib.write(wb, { bookType: 'biff8', type: 'array' }));
  const bytes = Array.prototype.map.call(u8, b => b > 127 ? b - 256 : b);   // Apps Script wants signed bytes

  const file = DriveApp.createFile(
    Utilities.newBlob(bytes, "application/vnd.ms-excel", tmpName + ".xls")
  );

  const fileId = file.getId();
  scheduleExportCleanup_(fileId);   // backup, in case the dialog is closed early

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:13px;padding:10px;">'
    + '<p>From <b>' + sheetName + '</b>: <b>' + (out.length - 1) + '</b> row(s) exported'
    + (skipped ? ', <b>' + skipped + '</b> hidden row(s) skipped' : '')
    + '.</p>'
    + '<p><a id="dl" href="' + file.getDownloadUrl() + '" target="_blank">Download '
    + tmpName + '.xls</a></p>'
    + '<p id="msg" style="color:#777;font-size:11px;">'
    + 'The file is removed from Drive 15 seconds after you click the link.</p>'
    + '<script>'
    + 'document.getElementById("dl").addEventListener("click", function(){'
    + '  var n = 15;'
    + '  var m = document.getElementById("msg");'
    + '  var t = setInterval(function(){'
    + '    n--; m.textContent = "Removing from Drive in " + n + "s...";'
    + '    if (n <= 0) {'
    + '      clearInterval(t);'
    + '      google.script.run'
    + '        .withSuccessHandler(function(){ m.textContent = "Removed from Drive."; })'
    + '        .withFailureHandler(function(e){ m.textContent = "Could not remove: " + e.message; })'
    + '        .deleteExportFile("' + fileId + '");'
    + '    }'
    + '  }, 1000);'
    + '});'
    + '<\/script>'
    + '</div>'
  ).setWidth(380).setHeight(170);

  ui.showModalDialog(html, "Export complete");
}


/** Called from the export dialog after its 15-second countdown. */
function deleteExportFile(fileId) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (e) {
    // already gone - fine
  }
  clearExportCleanupTriggers_(fileId);
  return true;
}


/** Backup cleanup: trashes the file even if the user closes the dialog. */
function scheduleExportCleanup_(fileId) {
  const trigger = ScriptApp.newTrigger('cleanupExportFile_')
    .timeBased()
    .after(60 * 1000)      // Apps Script won't fire reliably faster than ~1 min
    .create();
  PropertiesService.getScriptProperties()
    .setProperty('exportCleanup_' + trigger.getUniqueId(), fileId);
}


function cleanupExportFile_() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(function (key) {
    if (key.indexOf('exportCleanup_') !== 0) return;
    try { DriveApp.getFileById(all[key]).setTrashed(true); } catch (e) {}
    props.deleteProperty(key);
  });
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'cleanupExportFile_') ScriptApp.deleteTrigger(t);
  });
}


function clearExportCleanupTriggers_(fileId) {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(function (key) {
    if (key.indexOf('exportCleanup_') === 0 && all[key] === fileId) props.deleteProperty(key);
  });
}


/**
 * Returns an array of booleans, one per data row, saying whether it's visible.
 * Covers both filter-hidden and manually hidden rows.
 *
 * Uses the Sheets API to fetch all row metadata in ONE call. The obvious
 * isRowHiddenByFilter/isRowHiddenByUser loop costs two round trips per row,
 * which is fine for 25 rows and painful for 3000.
 */
function visibleRowFlags_(sheet, firstRow, nRows) {
  const flags = new Array(nRows).fill(true);

  try {
    const res = Sheets.Spreadsheets.get(
      sheet.getParent().getId(),
      {
        ranges: sheet.getName() + "!A" + firstRow + ":A" + (firstRow + nRows - 1),
        fields: "sheets/data/rowMetadata(hiddenByUser,hiddenByFilter)"
      }
    );

    const meta = res.sheets
      && res.sheets[0]
      && res.sheets[0].data
      && res.sheets[0].data[0]
      && res.sheets[0].data[0].rowMetadata;

    if (!meta) return flags;

    for (let i = 0; i < nRows && i < meta.length; i++) {
      if (meta[i].hiddenByUser || meta[i].hiddenByFilter) flags[i] = false;
    }
    return flags;

  } catch (e) {
    // Advanced Sheets Service not enabled, or the call failed — fall back
    // to the slow-but-always-works version.
    for (let i = 0; i < nRows; i++) {
      const r = firstRow + i;
      if (sheet.isRowHiddenByFilter(r) || sheet.isRowHiddenByUser(r)) flags[i] = false;
    }
    return flags;
  }
}


// "A" -> 0, "B" -> 1, "AA" -> 26
function colLetterToIndex_(letter) {
  const s = String(letter).trim().toUpperCase();
  if (!/^[A-Z]+$/.test(s)) return -1;
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n - 1;
}


/**
 * Renders one value as a CSV field.
 * Quotes anything containing a comma, quote, or newline; doubles inner quotes.
 * Whole numbers lose their trailing .0 so 85 exports as 85, not 85.0.
 */
function csvCell_(v) {
  if (v === null || v === undefined) return "";
  let s;
  if (typeof v === "number") {
    s = (v % 1 === 0) ? String(Math.round(v)) : String(v);
  } else if (v instanceof Date) {
    s = Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  } else {
    s = String(v);
  }
  if (/[",\r\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
  return s;
}
