/***** COPY TO IMPORT LIST — sticker rows on Print Console → Import List *****
 * Button "COPY TO IMPORT LIST" on Print Console (column A, row 3) and Product Tools menu.
 * Each product of the sticker list (row 8 down) is ADDED below the Import List's rows,
 * once; products already in the Import List (same barcode or same name) are skipped.
 * Import List columns: A Item Code | B Product Name | C Brand | D NPP | E MRP | F Shell Life | G Premium Rating
 */

const TIL_BUTTON_TITLE_ = 'COPY TO IMPORT LIST';
const TIL_BUTTON_PNG_ = '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'; // 328x152 (4x), shown at 82x38

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

// Puts the COPY TO IMPORT LIST button on Print Console (column A, row 3), or refreshes its picture.
function placeImportListButton(quiet) {
  const r = amPlaceButton_(TIL_BUTTON_TITLE_, TIL_BUTTON_PNG_, 1, 3, 8, 2, 'stickersToImportList');
  if (!quiet) SpreadsheetApp.getActiveSpreadsheet().toast('COPY TO IMPORT LIST button ' + r + '. You can drag it anywhere.', 'Buttons', 5);
}

// Menu: put both Print Console buttons (ADD MANY, COPY TO IMPORT LIST) in place, or refresh their pictures.
function placePrintConsoleButtons() {
  placeAddManyButton(true);
  placeImportListButton(true);
  SpreadsheetApp.getActiveSpreadsheet().toast('Buttons are on Print Console. You can drag them anywhere.', 'Buttons', 5);
}
