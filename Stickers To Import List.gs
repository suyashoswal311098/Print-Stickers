/***** COPY TO IMPORT LIST — sticker rows on Print Console → Import List *****
 * Button "COPY TO IMPORT LIST" on Print Console (column A, row 3) and Product Tools menu.
 * Each product of the sticker list (row 8 down) is ADDED below the Import List's rows,
 * once; products already in the Import List (same barcode or same name) are skipped.
 * Import List columns: A Item Code | B Product Name | C Brand | D NPP | E MRP | F Shell Life | G Premium Rating
 */

const TIL_BUTTON_TITLE_ = 'COPY TO IMPORT LIST';
const TIL_BUTTON_PNG_ = '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';

function stickersToImportList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const work = ss.getSheetByName('Print Console');
  const imp = ss.getSheetByName('Import List');
  const prod = ss.getSheetByName('Product List');
  if (!work || !imp || !prod) { ui.alert("Sheets 'Print Console', 'Import List' and 'Product List' must exist."); return; }

  // 1. Sticker rows (row 8 down): A code, B name, G brand, H NPP, I MRP
  const FIRST = 8;
  const nSt = work.getLastRow() - FIRST + 1;
  if (nSt <= 0) { ui.alert('There are no stickers in the list.'); return; }
  const st = work.getRange(FIRST, 1, nSt, 9).getValues();
  const stDisp = work.getRange(FIRST, 1, nSt, 1).getDisplayValues();

  // 2. Product List: shelf life (G) + premium rating (I) by barcode / name
  const nPr = Math.max(prod.getLastRow() - 1, 1);
  const pr = prod.getRange(2, 1, nPr, 9).getValues();
  const prBar = prod.getRange(2, 2, nPr, 1).getDisplayValues();
  const byBar = {}, byName = {};
  pr.forEach((r, i) => {
    const b = normBar_(prBar[i][0]), n = normName_(r[2]);
    if (b && !byBar[b]) byBar[b] = r;
    if (n && !byName[n]) byName[n] = r;
  });

  // 3. Import List rows already there (row 4 down, last row with A or B filled)
  const IMP_FIRST = 4;
  const impMax = imp.getMaxRows();
  let lastUsed = IMP_FIRST - 1;
  const have = { bar: {}, name: {} };
  if (impMax >= IMP_FIRST) {
    const ab = imp.getRange(IMP_FIRST, 1, impMax - IMP_FIRST + 1, 2).getDisplayValues();
    ab.forEach((r, i) => {
      const b = normBar_(r[0]), n = normName_(r[1]);
      if (b || n) lastUsed = IMP_FIRST + i;
      if (b) have.bar[b] = 1;
      if (n) have.name[n] = 1;
    });
  }

  // 4. New rows, one per product
  const out = [];
  let skipped = 0;
  st.forEach((r, i) => {
    const bar = String(stDisp[i][0]).trim();
    const name = String(r[1]).trim();
    if (!bar && !name) return;
    const b = normBar_(bar), n = normName_(name);
    if ((b && have.bar[b]) || (n && have.name[n])) { skipped++; return; }
    if (b) have.bar[b] = 1;
    if (n) have.name[n] = 1;
    const p = (b && byBar[b]) || (n && byName[n]) || null;
    out.push([bar, name, r[6], r[7], r[8], p ? p[6] : '', p ? p[8] : '']);
  });

  if (!out.length) {
    ui.alert('Nothing added: all ' + skipped + ' product(s) are already in the Import List.');
    return;
  }

  // 5. Write below the last used row (add sheet rows if needed); barcodes as text
  const start = lastUsed + 1;
  const need = start + out.length - 1;
  if (need > imp.getMaxRows()) imp.insertRowsAfter(imp.getMaxRows(), need - imp.getMaxRows());
  imp.getRange(start, 1, out.length, 1).setNumberFormat('@');
  imp.getRange(start, 1, out.length, 7).setValues(out);
  const rule = prod.getRange(2, 9).getDataValidation(); // Premium Rating dropdown, like Import Products
  if (rule) imp.getRange(start, 7, out.length, 1).setDataValidation(rule);

  ui.alert('Copy to Import List',
    'Added ' + out.length + ' product(s) to the Import List (rows ' + start + '–' + need + ').' +
    (skipped ? '\nSkipped ' + skipped + ' already there.' : ''),
    ui.ButtonSet.OK);
}

// Puts the COPY TO IMPORT LIST button on Print Console once (column A, row 3). Safe to run again.
function placeImportListButton(quiet) {
  const work = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Print Console');
  if (!work) return;
  if (work.getImages().some(img => img.getAltTextTitle() === TIL_BUTTON_TITLE_)) {
    if (!quiet) SpreadsheetApp.getActiveSpreadsheet().toast('The COPY TO IMPORT LIST button is already on Print Console.', 'Buttons', 5);
    return;
  }
  const blob = Utilities.newBlob(Utilities.base64Decode(TIL_BUTTON_PNG_), 'image/png', 'copy-to-import-list.png');
  const img = work.insertImage(blob, 1, 3, 8, 2); // column A, row 3
  img.setWidth(86).setHeight(40);
  img.setAltTextTitle(TIL_BUTTON_TITLE_);
  img.assignScript('stickersToImportList');
  if (!quiet) SpreadsheetApp.getActiveSpreadsheet().toast('COPY TO IMPORT LIST button placed. You can drag it anywhere.', 'Buttons', 5);
}

// Menu: put both Print Console buttons (ADD MANY, COPY TO IMPORT LIST) in place if missing.
function placePrintConsoleButtons() {
  placeAddManyButton(true);
  placeImportListButton(true);
  SpreadsheetApp.getActiveSpreadsheet().toast('Buttons are on Print Console. You can drag them anywhere.', 'Buttons', 5);
}
