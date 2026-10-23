/***** PURCHASE WINDOW — LOOSE (buy by weight, split into the Product List's pack sizes) or PACKED (scan barcode /
 * type name, packs + rate per pack); both put the lines in "Purchase CSV" *****
 * Product Tools → Purchase CSV for POS → "Add loose purchase…" opens "Loose Purchase Dialog".
 * The window does the thinking (find the pack sizes in the product names, split the kg, tax calculator).
 * This file only (1) hands the window the Product List and (2) writes the finished lines into the sheet:
 *   A = dropdown text "barcode name", D = packs, F = tax code, K = Cost override (cost per pack INCLUDING tax; the window
 *   works ex tax and the server adds the tax with pcInclCost_).
 * B / C / I / J / L are the header-row formulas of the sheet and are never written here.
 */

function openLoosePurchase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(PC_SHEET_)) { SpreadsheetApp.getUi().alert('Run Product Tools → Purchase CSV for POS → Set up / repair sheet first.'); return; }
  pcEnsureLayout_();   // old sheet formulas → repaired first (lines kept)
  const html = HtmlService.createHtmlOutputFromFile('Loose Purchase Dialog').setWidth(1000).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Purchase - loose (kg) or packed (pcs)');
}

// Product List for the window: [[barcode, name, NPP, dropdown text, MRP, shelf days], …] + tax codes + default tax.
function lpGetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const list = ss.getSheetByName('Product List');
  const n = Math.max(list.getLastRow() - 1, 1);
  const bn = list.getRange(2, 2, n, 2).getDisplayValues();
  const npp = list.getRange(2, 5, n, 3).getValues();   // E NPP, F MRP, G shelf life (days)

  // dropdown text ("barcode name") of the Print Console B2 list, by barcode
  const dd = {};
  const rule = ss.getSheetByName('Print Console').getRange('B2').getDataValidation();
  if (rule) {
    const type = rule.getCriteriaType(), vals = rule.getCriteriaValues();
    let names = [];
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) names = vals[0].getValues().flat();
    else if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) names = vals[0];
    names.forEach(v => {
      const t = String(v).trim(), m = t.match(/^(\S+)\s/);
      const k = m && normBar_(m[1]);
      if (k && !dd[k]) dd[k] = t;
    });
  }
  const products = [];
  bn.forEach((r, i) => {
    const code = String(r[0]).trim(), name = String(r[1]).trim();
    if (!code || !name) return;
    products.push([code, name, npp[i][0] === '' ? '' : Number(npp[i][0]), dd[normBar_(code)] || (code + ' ' + name), npp[i][1] === '' ? '' : Number(npp[i][1]), Number(npp[i][2]) || '']);
  });
  const sh = ss.getSheetByName(PC_SHEET_);
  return {
    products: products,
    taxes: PC_TAX_CODES_.map(t => [t[1], t[2]]),
    defTax: String(sh.getRange('B2').getValue() || '').trim()
  };
}

// lines = [{dd, packs, cost, tax}] → appended below the last product in column A. Returns {first, last}.
function lpWrite(lines) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PC_SHEET_);
  if (!sh) throw new Error('Run Set up / repair sheet first.');
  if (!lines || !lines.length) throw new Error('No lines to put in the sheet.');
  const lastRow = PC_FIRST_ + PC_ROWS_ - 1;
  const a = sh.getRange(PC_FIRST_, 1, PC_ROWS_, 1).getValues();
  let used = 0;
  a.forEach((r, i) => { if (String(r[0]).trim()) used = i + 1; });
  const first = PC_FIRST_ + used, last = first + lines.length - 1;
  if (last > lastRow) throw new Error('The sheet has room for ' + PC_ROWS_ + ' lines only (' + used + ' used).');
  const taxOk = pcTaxCodes_();
  lines.forEach(l => {
    if (!l.dd || !(l.packs > 0) || !(l.cost >= 0)) throw new Error('A line is incomplete: ' + JSON.stringify(l));
    if (l.tax && taxOk.indexOf(l.tax) < 0) throw new Error('Tax code "' + l.tax + '" is not a POS code.');
  });
  const n = lines.length;
  sh.getRange(first, 1, n, 1).setValues(lines.map(l => [l.dd]));
  sh.getRange(first, 4, n, 1).setValues(lines.map(l => [l.packs]));
  sh.getRange(first, 6, n, 1).setValues(lines.map(l => [l.tax || '']));
  sh.getRange(first, 11, n, 1).setValues(lines.map(l => [pcInclCost_(l.cost, l.tax)]));
  // H = Expiry (text MM/dd/yyyy, today + shelf life, worked out in the window); blank when no shelf life
  sh.getRange(first, 8, n, 1).setNumberFormat('@').setValues(lines.map(l => [/^\d{2}\/\d{2}\/\d{4}$/.test(String(l.exp || '')) ? l.exp : '']));   // K = cost per pack INCLUDING tax (the window works ex tax)
  SpreadsheetApp.flush();
  return { first: first, last: last };
}

// "Download" button of the window: same checks + the same two files as the DOWNLOAD XLSX FILE button.
function lpDownload() { return pcBuildFiles(); }

// ✎ in the window: change packs / cost / tax of a line already in the sheet (row must still hold the same product).
function lpUpdate(row, dd, v) {
  const sh = lpRowCheck_(row, dd);
  if (!(v.packs > 0) || !(v.cost >= 0)) throw new Error('Packs and cost are not valid.');
  if (v.tax && pcTaxCodes_().indexOf(v.tax) < 0) throw new Error('Tax code "' + v.tax + '" is not a POS code.');
  sh.getRange(row, 4).setValue(v.packs);
  sh.getRange(row, 6).setValue(v.tax || '');
  sh.getRange(row, 11).setValue(pcInclCost_(v.cost, v.tax));
  SpreadsheetApp.flush();
  return true;
}

// ✕ on a line already in the sheet: empties that row's typed cells (the formulas stay; rows below do not move).
function lpClear(row, dd) {
  const sh = lpRowCheck_(row, dd);
  [1, 4, 5, 6, 7, 8, 11].forEach(c => sh.getRange(row, c).clearContent());
  SpreadsheetApp.flush();
  return true;
}

function lpRowCheck_(row, dd) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PC_SHEET_);
  if (!sh) throw new Error('The Purchase CSV sheet is missing.');
  if (row < PC_FIRST_ || row > PC_FIRST_ + PC_ROWS_ - 1) throw new Error('Row ' + row + ' is outside the purchase lines.');
  if (String(sh.getRange(row, 1).getValue()).trim() !== String(dd).trim()) throw new Error('Row ' + row + ' in the sheet is not this product any more - change it in the sheet.');
  return sh;
}

// NPP / MRP typed in the window (user's choice: also saved in Product List). changes = [{code, npp?, mrp?}];
// only the given fields are written (E = NPP, F = MRP) in the first Product List row with that barcode. Returns rows changed.
function lpSavePrices(changes) {
  const list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Product List');
  const n = Math.max(list.getLastRow() - 1, 1);
  const bars = list.getRange(2, 2, n, 1).getDisplayValues();
  const at = {};
  bars.forEach((r, i) => { const k = normBar_(r[0]); if (k && at[k] === undefined) at[k] = i + 2; });
  let done = 0;
  (changes || []).forEach(c => {
    const row = at[normBar_(c.code)];
    if (!row) return;
    const ok = v => typeof v === 'number' && isFinite(v) && v >= 0;
    if (ok(c.npp)) list.getRange(row, 5).setValue(c.npp);
    if (ok(c.mrp)) list.getRange(row, 6).setValue(c.mrp);
    if (ok(c.npp) || ok(c.mrp)) done++;
  });
  return done;
}
