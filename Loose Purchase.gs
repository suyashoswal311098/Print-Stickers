/***** LOOSE PURCHASE — buy by weight, split into the Product List's pack sizes, put the lines in "Purchase CSV" *****
 * Product Tools → Purchase CSV for POS → "Add loose purchase…" opens "Loose Purchase Dialog".
 * The window does the thinking (find the pack sizes in the product names, split the kg, tax calculator).
 * This file only (1) hands the window the Product List and (2) writes the finished lines into the sheet:
 *   A = dropdown text "barcode name", D = packs, F = tax code, K = Cost override (cost per pack, tax excluded).
 * B / C / I / J / L are the header-row formulas of the sheet and are never written here.
 */

function openLoosePurchase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(PC_SHEET_)) { SpreadsheetApp.getUi().alert('Run Product Tools → Purchase CSV for POS → Set up / repair sheet first.'); return; }
  const html = HtmlService.createHtmlOutputFromFile('Loose Purchase Dialog').setWidth(900).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Loose purchase');
}

// Product List for the window: [[barcode, name, NPP, dropdown text], …] + tax codes + default tax.
function lpGetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const list = ss.getSheetByName('Product List');
  const n = Math.max(list.getLastRow() - 1, 1);
  const bn = list.getRange(2, 2, n, 2).getDisplayValues();
  const npp = list.getRange(2, 5, n, 1).getValues();

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
    products.push([code, name, npp[i][0] === '' ? '' : Number(npp[i][0]), dd[normBar_(code)] || (code + ' ' + name)]);
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
  sh.getRange(first, 11, n, 1).setValues(lines.map(l => [l.cost]));
  SpreadsheetApp.flush();
  return { first: first, last: last };
}

// "Download XLSX" button of the window: same checks + file as the menu item.
function lpDownload() { return pcMakeFile_(); }
