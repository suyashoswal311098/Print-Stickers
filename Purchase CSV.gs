/***** PURCHASE CSV — build the POS "Add Purchase by CSV" file from the Product List *****
 * Sheet "Purchase CSV": pick a product in column A (dropdown = the Print Console B2 list, "barcode name",
 * type to search). FORMULAS IN THE HEADER ROW (no script on edit) fill Product Code (the barcode in front
 * of the picked text, so leading zeros stay) + Net Unit Cost (the Cost override in column N; NPP is the selling
 * price and is NOT used as cost). You type Quantity, optionally Variant / Tax Rate Code / Discount / Expiry.
 * Tax Rate Code left empty = the Default Tax Rate Code in B2 (applied in the file).
 * Menu Product Tools → Purchase CSV for POS → Download XLSX gives a file with exactly the POS columns:
 *   Product Code | Net Unit Cost | Quantity | Product Variant | Tax Rate Code | Discount | Expiry
 * Sheet columns: A Product | B..H = the 7 POS columns | I Margin vs NPP (info, "308.50 (56.1%)") | J NPP, K MRP (Product
 *   List values; typing a new one saves it in Product List - Purchase NPP MRP.gs) | L Total incl tax (info) | M Check (info)
 *   | N Cost override INCL tax (optional) | O Tax used (info) | P Rate ex tax (info) | Q Total ex tax (info)
 */

const PC_SHEET_ = 'Purchase CSV';
const PC_TAX_SHEET_ = 'POS Tax Codes';
const PC_HEAD_ROW_ = 4;
const PC_FIRST_ = 5;
const PC_ROWS_ = 500;
const PC_HEADERS_ = ['Product Code', 'Net Unit Cost', 'Quantity', 'Product Variant', 'Tax Rate Code', 'Discount', 'Expiry'];
// POS "TAX RATE CODE" window (Name, Code, Tax Rate, Type). The CODE is what goes in the file.
const PC_TAX_CODES_ = [
  ['0%GST', '0%GST', 0, 1],
  ['18%GST', '18%GST', 18, 1],
  ['05%GST', '05%GST', 5, 1],
  ['5%GST', '12%GST', 5, 1],
  ['5% GST', '28%GST', 5, 1]
];
const PC_EXPIRY_FORMAT_ = 'MM/dd/yyyy';   // as in the POS sample file (09/26/2019); a date typed in Expiry is written like this
const PC_COST_COL_ = 14;   // N = Cost override INCL tax (typed, or written by the purchase window)
const PC_SPILL_COLS_ = [2, 3, 9, 12, 13, 15, 16, 17];   // filled by the header-row formulas: must stay empty below
const PC_XLSX_MIME_ = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// Cost × (1 + tax %) for a Tax Rate Code (POS Tax Codes table); no / unknown code = no tax. 4 decimals.
function pcInclCost_(cost, taxCode) {
  const t = PC_TAX_CODES_.filter(x => x[1] === taxCode)[0];
  return Math.round(cost * (1 + (t ? t[2] : 0) / 100) * 10000) / 10000;
}

// Layout version of the Purchase CSV sheet. Raise it whenever setupPurchaseSheet changes formulas/columns:
// the LOOSE PURCHASE / DOWNLOAD buttons then repair the sheet by themselves (typed lines are kept).
const PC_LAYOUT_VERSION_ = '11';   // 11 = GST breakup block S:W
// (previous:)   // 10 = Net Unit Cost no longer from NPP (NPP = selling price); 9 = I Margin vs NPP, NPP / MRP J / K, info L–Q; 8 = Net Unit Cost (C) coloured + typing there = Cost override; 4 = Net Unit Cost incl tax; 5 = ADD PURCHASE button picture; 6 = O NPP / P MRP; 7 = NPP / MRP moved to I / J (info columns now K–P)
function pcEnsureLayout_() {
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty('pc_layout') === PC_LAYOUT_VERSION_) return false;
  setupPurchaseSheet();
  return true;
}

function pcTaxCodes_() { return PC_TAX_CODES_.map(t => t[1]); }

// Creates (or repairs, without clearing typed lines) the Purchase CSV + POS Tax Codes sheets.
function setupPurchaseSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const work = ss.getSheetByName('Print Console');
  const rule = work && work.getRange('B2').getDataValidation();
  if (!rule) { ui.alert("Print Console B2 has no product dropdown, so the Product dropdown cannot be copied."); return; }

  // POS tax code reference
  let tx = ss.getSheetByName(PC_TAX_SHEET_) || ss.insertSheet(PC_TAX_SHEET_);
  tx.clear();
  tx.getRange(1, 1, 1, 4).setValues([['Name', 'Code', 'Tax Rate', 'Type']])
    .setBackground('#4a86c8').setFontColor('#ffffff').setFontWeight('bold');
  tx.getRange(2, 1, PC_TAX_CODES_.length, 4).setNumberFormat('@').setValues(PC_TAX_CODES_.map(r => r.map(String)));
  tx.getRange(2, 3, PC_TAX_CODES_.length, 1).setNumberFormat('0.0000');
  tx.getRange(2, 3, PC_TAX_CODES_.length, 1).setValues(PC_TAX_CODES_.map(r => [r[2]]));
  tx.getRange(8, 1).setValue('Copy of the POS "TAX RATE CODE" window. The file uses the CODE column. ' +
    'Note: in the POS, code 12%GST has rate 5 and code 28%GST has rate 5 - fix them in the POS if that is not intended.')
    .setFontColor('#a33').setWrap(false);
  tx.setColumnWidths(1, 4, 110);

  // Purchase sheet
  const isNew = !ss.getSheetByName(PC_SHEET_);
  const sh = ss.getSheetByName(PC_SHEET_) || ss.insertSheet(PC_SHEET_);
  const lastRow = PC_FIRST_ + PC_ROWS_ - 1;
  if (sh.getMaxRows() < lastRow) sh.insertRowsAfter(sh.getMaxRows(), lastRow - sh.getMaxRows());
  // I = Margin, J / K = NPP / MRP right after Expiry. Older sheets get there by moving / inserting WHOLE columns, so typed
  // values (NPP / MRP, Cost override) travel with them: layout 7–8 (NPP in I) → 1 column in front; layout 6 (NPP in O) →
  // O:P moved to I, then 1 in front; layout ≤ 5 (no NPP) → 3 new columns after H.
  const headAt = c => c <= sh.getMaxColumns() ? String(sh.getRange(PC_HEAD_ROW_, c).getValue()) : '';
  if (!isNew && headAt(PC_NPP_COL_).indexOf('NPP') !== 0) {
    if (headAt(9).indexOf('NPP') === 0) sh.insertColumnsAfter(8, 1);
    else if (headAt(15).indexOf('NPP') === 0) { sh.moveColumns(sh.getRange('O:P'), 9); sh.insertColumnsAfter(8, 1); }
    else sh.insertColumnsAfter(8, 3);
  }
  if (sh.getMaxColumns() < 17) sh.insertColumnsAfter(sh.getMaxColumns(), 17 - sh.getMaxColumns());

  sh.getRange('A1').setValue('PURCHASE FILE FOR POS  (Add Purchase by CSV)').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setValue('Default Tax Rate Code  →').setHorizontalAlignment('right').setFontWeight('bold');
  sh.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(pcTaxCodes_(), true).setAllowInvalid(false).build()).setBackground('#fff2cc');
  sh.getRange('G2').setValue('Total incl tax').setHorizontalAlignment('right').setFontWeight('bold');
  sh.getRange('I2:Q2').clearContent();
  sh.getRange('H2').setFormula('=SUM(L' + PC_FIRST_ + ':L' + lastRow + ')').setNumberFormat('#,##0.00').setFontWeight('bold');
  sh.getRange('P2').setValue('Total ex tax').setHorizontalAlignment('right').setFontWeight('bold');
  sh.getRange('Q2').setFormula('=SUM(Q' + PC_FIRST_ + ':Q' + lastRow + ')').setNumberFormat('#,##0.00').setFontWeight('bold');
  sh.getRange('A3').setValue('Pick/type a product in column A and type Quantity. B, C, I, L, M, O, P, Q are formulas in the header row (do not type there). ' +
    'B–H go into the file in the POS order (Net Unit Cost is the rate INCLUDING tax). C Net Unit Cost: type a new cost (incl tax) to override it. I Margin = NPP − Net Unit Cost (% of NPP). J NPP / K MRP: type a new value to change it in Product List. The grey columns are only for you.').setFontColor('#666');

  // Header row: plain labels, except B, C, I, L, M, O, P, Q = label + ONE array formula that fills the whole column.
  // Spill cells below must be empty, so old values/formulas in those columns are cleared first.
  PC_SPILL_COLS_.forEach(c => sh.getRange(PC_FIRST_, c, PC_ROWS_, 1).clearContent());
  const A = 'A' + PC_FIRST_ + ':A' + lastRow, B = 'B' + PC_FIRST_ + ':B' + lastRow, C = 'C' + PC_FIRST_ + ':C' + lastRow;
  const D = 'D' + PC_FIRST_ + ':D' + lastRow, K = 'N' + PC_FIRST_ + ':N' + lastRow, L = 'O' + PC_FIRST_ + ':O' + lastRow;   // K = cost override, L = tax used (variable names kept)
  const F = 'F' + PC_FIRST_ + ':F' + lastRow, NPP = 'J' + PC_FIRST_ + ':J' + lastRow;
  const bars = "'Product List'!$B$2:$B&\"\"";
  // tax % of the tax code used on the line (no / unknown code = 0)
  const taxPct = 'IFERROR(VLOOKUP(' + L + ",'" + PC_TAX_SHEET_ + "'!$B$2:$C$" + (PC_TAX_CODES_.length + 1) + ',2,FALSE),0)';
  sh.getRange(PC_HEAD_ROW_, 1, 1, 17).setNumberFormat('General');
  sh.getRange(PC_HEAD_ROW_, 1, 1, 17).setValues([['Product (barcode + name)'].concat(PC_HEADERS_, ['', 'NPP ✎ (Product List)', 'MRP ✎ (Product List)', '', '', 'Cost override INCL tax (optional)', '', '', ''])]);
  sh.getRange(PC_HEAD_ROW_, 2).setFormula('={"Product Code";ARRAYFORMULA(IF(' + A + '="","",IFERROR(REGEXEXTRACT(' + A + '&"","^\\S+"),"")))}');
  // Net Unit Cost = price INCLUDING tax = the Cost override (column N; typed in C, or written by the purchase window).
  // NPP is the SELLING price (user, 24 Oct), so it is never used as the cost: no override = empty + "⚠ type Net Unit Cost".
  sh.getRange(PC_HEAD_ROW_, 3).setFormula('={"Net Unit Cost";ARRAYFORMULA(IF(' + B + '="","",IF(ISNUMBER(' + K + '),' + K + ',"")))}');
  // I Margin vs NPP: amount and % of NPP in one cell, e.g. NPP 550, cost 241.50 → "308.50  (56.1%)"
  sh.getRange(PC_HEAD_ROW_, 9).setFormula('={"Margin vs NPP (info)";ARRAYFORMULA(IF(ISNUMBER(' + C + ')*ISNUMBER(' + NPP + ')*(' + NPP + '>0),TEXT(' + NPP + '-' + C +
    ',"0.00")&"  ("&TEXT((' + NPP + '-' + C + ')/' + NPP + '*100,"0.0")&"%)",""))}');
  sh.getRange(PC_HEAD_ROW_, 12).setFormula('={"Total incl tax (info)";ARRAYFORMULA(IF(ISNUMBER(' + C + ')*ISNUMBER(' + D + '),ROUND(' + C + '*' + D + ',2),""))}');
  sh.getRange(PC_HEAD_ROW_, 13).setFormula('={"Check (info)";ARRAYFORMULA(IF(' + A + '="","",IF(ISNA(VLOOKUP(' + B + ',{' + bars + ',' + bars + '},2,FALSE)),"⚠ not in Product List",' +
    'IF(NOT(ISNUMBER(' + D + ')),"⚠ type Quantity",IF(' + D + '<=0,"⚠ Quantity must be above 0",IF(NOT(ISNUMBER(' + C + ')),"⚠ type Net Unit Cost (column C)","✔ ok"))))))}');
  sh.getRange(PC_HEAD_ROW_, 1, 1, 17).setBackground('#4a86c8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(PC_HEAD_ROW_, 1).setBackground('#1f3d66');
  sh.getRange(PC_HEAD_ROW_, 15).setFormula('={"Tax used in file (info)";ARRAYFORMULA(IF(' + A + '="","",IF(' + F + '<>"",' + F + ',IF($B$2="","⚠ no tax",$B$2))))}');
  sh.getRange(PC_HEAD_ROW_, 16).setFormula('={"Rate ex tax (info)";ARRAYFORMULA(IF(ISNUMBER(' + C + '),ROUND(' + C + '/(1+' + taxPct + '/100),2),""))}');
  sh.getRange(PC_HEAD_ROW_, 17).setFormula('={"Total ex tax (info)";ARRAYFORMULA(IF(ISNUMBER(' + C + ')*ISNUMBER(' + D + '),ROUND(' + C + '*' + D + '/(1+' + taxPct + '/100),2),""))}');
  sh.getRange(PC_HEAD_ROW_, 12, 1, 6).setBackground('#7f7f7f');
  sh.getRange(PC_HEAD_ROW_, 9).setBackground('#6a1b9a');   // I Margin: purple
  // J / K: NPP / MRP from Product List, typed changes are saved back there (onEdit → pcOnEdit_)
  sh.getRange(PC_HEAD_ROW_, PC_NPP_COL_, 1, 2).setBackground('#b45f06');
  sh.getRange(PC_HEAD_ROW_, 3).setBackground('#137333');   // C Net Unit Cost: green (typing there = Cost override, pcOnEdit_)
  // GST breakup (S:W, rows 1-8): per tax code taxable value, CGST, SGST, total tax - same figures as a supplier bill's tax table.
  // Taxable = sum of "Total ex tax" (Q) of the lines using the code, tax = sum of "Total incl tax" (L) minus that.
  if (sh.getMaxColumns() < 23) sh.insertColumnsAfter(sh.getMaxColumns(), 23 - sh.getMaxColumns());
  const nT = PC_TAX_CODES_.length, tEnd = 3 + nT - 1, totR = tEnd + 1;
  sh.getRange(1, 19, totR, 5).clearContent();
  sh.getRange('S1').setValue('GST BREAKUP (info)').setFontWeight('bold');
  sh.getRange(2, 19, 1, 5).setValues([['Tax code', 'Taxable value', 'CGST', 'SGST', 'Total tax']]).setBackground('#7f7f7f').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  const rg = (c) => '$' + c + '$' + PC_FIRST_ + ':$' + c + '$' + lastRow;
  for (let i = 0; i < nT; i++) {
    const r = 3 + i;
    sh.getRange(r, 19).setNumberFormat('@').setValue(String(PC_TAX_CODES_[i][1]));
    sh.getRange(r, 20).setFormula('=ROUND(SUMIF(' + rg('O') + ',S' + r + ',' + rg('Q') + '),2)');
    sh.getRange(r, 23).setFormula('=ROUND(SUMIF(' + rg('O') + ',S' + r + ',' + rg('L') + ')-T' + r + ',2)');
    sh.getRange(r, 21).setFormula('=ROUND(W' + r + '/2,2)');
    sh.getRange(r, 22).setFormula('=W' + r + '-U' + r);
  }
  sh.getRange(totR, 19).setValue('Total').setFontWeight('bold');
  ['T', 'U', 'V', 'W'].forEach(c => sh.getRange(c + totR).setFormula('=SUM(' + c + '3:' + c + tEnd + ')').setFontWeight('bold'));
  sh.getRange(3, 20, totR - 2, 4).setNumberFormat('#,##0.00').setFontColor('#444');
  sh.getRange(totR + 1, 19).setValue('Total incl tax = Taxable + Total tax (the supplier bill may differ by its round off)').setFontColor('#666').setFontSize(9);
  sh.setColumnWidths(19, 5, 100);
  sh.setFrozenRows(PC_HEAD_ROW_);

  // formats: variant + discount + expiry as plain text ("5%" stays as typed); code/cost come from the formulas
  sh.getRange(PC_FIRST_, 2, PC_ROWS_, 1).setNumberFormat('General');
  sh.getRange(PC_FIRST_, 5, PC_ROWS_, 1).setNumberFormat('@');
  sh.getRange(PC_FIRST_, 7, PC_ROWS_, 1).setNumberFormat('@');
  sh.getRange(PC_FIRST_, 3, PC_ROWS_, 1).setNumberFormat('0.00');
  sh.getRange(PC_FIRST_, 4, PC_ROWS_, 1).setNumberFormat('0.##');
  sh.getRange(PC_FIRST_, 8, PC_ROWS_, 1).setNumberFormat('@');
  sh.getRange(PC_FIRST_, 12, PC_ROWS_, 1).setNumberFormat('#,##0.00');
  sh.getRange(PC_FIRST_, 9, PC_ROWS_, 1).setBackground('#f3e5f5').setHorizontalAlignment('right').setFontWeight('bold').setFontColor('#000000');
  sh.getRange(PC_FIRST_, PC_COST_COL_, PC_ROWS_, 1).setNumberFormat('0.00');
  sh.getRange(PC_FIRST_, PC_NPP_COL_, PC_ROWS_, 2).setNumberFormat('0.00').setBackground('#fff8e1').setHorizontalAlignment('right').setFontColor('#000000');
  sh.getRange(PC_FIRST_, 3, PC_ROWS_, 1).setBackground('#e6f4ea').setFontWeight('bold');

  // dropdowns
  sh.getRange(PC_FIRST_, 1, PC_ROWS_, 1).setDataValidation(rule);
  sh.getRange(PC_FIRST_, 6, PC_ROWS_, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(pcTaxCodes_(), true).setAllowInvalid(false).build());

  sh.getRange(PC_FIRST_, 12, PC_ROWS_, 6).setFontColor('#666');
  sh.getRange(PC_FIRST_, 16, PC_ROWS_, 2).setNumberFormat('#,##0.00');
  sh.getRange(PC_FIRST_, 2, PC_ROWS_, 7).setHorizontalAlignment('left');

  const jr = sh.getRange(PC_FIRST_, 13, PC_ROWS_, 1), mr = sh.getRange(PC_FIRST_, 9, PC_ROWS_, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('-').setFontColor('#cc0000').setRanges([mr]).build(),   // loss
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('⚠').setFontColor('#cc0000').setBold(true).setRanges([jr]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('✔').setFontColor('#1e8e3e').setRanges([jr]).build()
  ]);

  sh.setColumnWidth(1, 380);
  [110, 95, 75, 100, 110, 80, 95, 150, 130, 130, 95, 230, 175, 170, 150, 160].forEach((w, i) => sh.setColumnWidth(i + 2, w));
  // NPP / MRP of the lines already there
  const typed = sh.getRange(PC_FIRST_, 1, PC_ROWS_, 1).getValues();
  let usedTo = 0;
  typed.forEach((r, i) => { if (String(r[0]).trim()) usedTo = i + 1; });
  pcFillPrices_(sh, PC_FIRST_, PC_FIRST_ + usedTo - 1);
  sh.activate();
  sh.getRange(PC_FIRST_, 1).activate();
  try { pcPlaceButtons(true); } catch (err) { console.warn('Purchase CSV buttons: ' + err.message); }
  PropertiesService.getDocumentProperties().setProperty('pc_layout', PC_LAYOUT_VERSION_);
  ss.toast(isNew ? 'Purchase CSV sheet created.' : 'Purchase CSV sheet checked / repaired (your lines were kept).', 'Purchase CSV', 5);
}

// Product List lookup: dropdown text ("<barcode> <name>", plain name, or plain barcode) → {code, name, cost}.
// Code = Product List column B as displayed (leading zeros kept), cost = NPP (column E).
function pcIndex_() {
  const list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Product List');
  const n = Math.max(list.getLastRow() - 1, 1);
  const bn = list.getRange(2, 2, n, 2).getDisplayValues();
  const npp = list.getRange(2, 5, n, 1).getValues();
  const sq = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toUpperCase();
  const byName = {}, byBar = {};
  bn.forEach((r, i) => {
    const v = { code: String(r[0]).trim(), name: sq(r[1]), cost: npp[i][0] === '' ? '' : Number(npp[i][0]) };
    if (v.name && !byName[v.name]) byName[v.name] = v;
    const bk = normBar_(r[0]);
    if (bk && !byBar[bk]) byBar[bk] = v;
  });
  return function lookup(text) {
    const full = sq(text);
    if (!full) return null;
    const m = full.match(/^(\S+)\s+(.+)$/);
    const p = m && byBar[normBar_(m[1])];
    if (p && (!p.name || p.name === m[2])) return p;
    return byName[full] || byBar[normBar_(full)] || null;
  };
}

function pcNum_(v) {
  if (typeof v === 'number') return v;
  const s = String(v == null ? '' : v).trim();
  return s !== '' && isFinite(s) ? Number(s) : NaN;
}

// Reads the sheet → {rows: [{code,cost,qty,variant,tax,disc,exp,row}], problems: ['Row 7: …']}
function pcCollect_() {
  SpreadsheetApp.flush();   // formulas in the header row must be calculated before reading B / C
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PC_SHEET_);
  if (!sh) return { rows: [], problems: ['The "' + PC_SHEET_ + '" sheet does not exist - run Set up / repair sheet first.'] };
  const n = sh.getLastRow() - PC_FIRST_ + 1;
  if (n <= 0) return { rows: [], problems: [] };
  const v = sh.getRange(PC_FIRST_, 1, n, 8).getValues();
  const code = sh.getRange(PC_FIRST_, 2, n, 1).getDisplayValues();
  const tz = Session.getScriptTimeZone();
  const taxOk = pcTaxCodes_();
  const defTax = String(sh.getRange('B2').getValue() || '').trim();
  let lookup = null;
  const rows = [], problems = [];
  v.forEach((r, i) => {
    const row = PC_FIRST_ + i;
    const prod = String(r[0]).trim();
    let c = String(code[i][0]).trim();
    if (!prod && r[3] === '') return;
    if (c) { lookup = lookup || pcIndex_(); const p = lookup(c); if (p) c = p.code; else c = ''; }   // code must exist in Product List
    const cost = pcNum_(r[2]), qty = pcNum_(r[3]);
    const tax = String(r[5]).trim() || defTax;
    const bad = [];
    if (!c) bad.push(prod ? 'product not found in Product List' : 'no product');
    if (isNaN(qty) || qty <= 0) bad.push('Quantity must be a number above 0');
    if (isNaN(cost) || cost < 0) bad.push('Net Unit Cost is empty - type the cost (incl tax) in column C');
    if (tax && taxOk.indexOf(tax) < 0) bad.push('Tax Rate Code "' + tax + '" is not a POS code');
    if (bad.length) { problems.push('Row ' + row + ': ' + bad.join('; ')); return; }
    const exp = r[7] instanceof Date ? Utilities.formatDate(r[7], tz, PC_EXPIRY_FORMAT_) : String(r[7]).trim();
    rows.push({ row: row, code: c, cost: cost, qty: qty, variant: String(r[4]).trim(), tax: tax, disc: String(r[6]).trim(), exp: exp });
  });
  return { rows: rows, problems: problems };
}

// Array of arrays in the POS column order; blanks are null; Discount stays text if it has a % sign.
function pcAoa_(rows) {
  const out = [PC_HEADERS_.slice()];
  rows.forEach(r => {
    const d = r.disc === '' ? null : (/^\d+(\.\d+)?$/.test(r.disc) ? Number(r.disc) : r.disc);
    out.push([r.code, r.cost, r.qty, r.variant || null, r.tax || null, d, r.exp || null]);
  });
  return out;
}

// Real .xlsx: Product Code / Variant are text cells (leading zeros kept). Returns signed bytes for Apps Script.
function pcXlsxBytes_(aoa, sheetName, textCols) {
  const lib = sheetJs_();
  const ws = lib.utils.aoa_to_sheet(aoa);
  for (let r = 2; r <= aoa.length; r++) {
    (textCols || ['A', 'D', 'E', 'G']).forEach(col => {
      const c = ws[col + r];
      if (c) { c.t = 's'; c.v = String(c.v); c.z = '@'; delete c.w; }
    });
  }
  const wb = lib.utils.book_new();
  lib.utils.book_append_sheet(wb, ws, sheetName || 'sample_purchase_products');   // default = same sheet name as the POS sample
  const u8 = new Uint8Array(lib.write(wb, { bookType: 'xlsx', type: 'array', bookSST: true }));
  return Array.prototype.map.call(u8, b => b > 127 ? b - 256 : b);
}

// .xls (Excel 97-2003, biff8) bytes exactly like exportProductsToXlsx: blanks = empty cells, codes are strings.
function pcXlsBytes_(aoa) {
  const lib = sheetJs_();
  const wb = lib.utils.book_new();
  lib.utils.book_append_sheet(wb, lib.utils.aoa_to_sheet(aoa), 'Worksheet');
  const u8 = new Uint8Array(lib.write(wb, { bookType: 'biff8', type: 'array' }));
  return Array.prototype.map.call(u8, b => b > 127 ? b - 256 : b);
}

// Product import sheet (same columns as Export Products to XLS: Code, Product Name, Article Code, Price, MRP,
// Variants_Name, Variants_Price) for the products on the purchase lines. pl = {bar, name, npp, mrp} parallel
// arrays of the Product List (bar as displayed). One row per product, in the order of the lines.
function pcProductAoa_(rows, pl) {
  const at = {};
  pl.bar.forEach((b, i) => { const k = normBar_(b); if (k && at[k] === undefined) at[k] = i; });
  const out = [['Code', 'Product Name', 'Article Code', 'Price', 'MRP', 'Variants_Name', 'Variants_Price']];
  const seen = {};
  rows.forEach(r => {
    const k = normBar_(r.code);
    if (seen[k]) return;
    seen[k] = true;
    const i = at[k];
    if (i === undefined) return;
    out.push([String(pl.bar[i]).trim(), pl.name[i], null, pl.npp[i] === '' ? null : pl.npp[i], pl.mrp[i] === '' ? null : pl.mrp[i], null, null]);
  });
  return out;
}

// (not used: the POS rejects .csv) UTF-8 CSV, no BOM, comma separated.
function pcCsv_(aoa) {
  const cell = v => {
    if (v === null || v === undefined) return '';
    const s = String(v);
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  return aoa.map(r => r.map(cell).join(',')).join('\r\n') + '\r\n';
}

function pcCheck() {
  const ui = SpreadsheetApp.getUi();
  const res = pcCollect_();
  if (res.problems.length) {
    ui.alert('Purchase CSV: please fix', res.problems.slice(0, 15).join('\n') +
      (res.problems.length > 15 ? '\n…and ' + (res.problems.length - 15) + ' more' : ''), ui.ButtonSet.OK);
  } else if (!res.rows.length) {
    ui.alert('Purchase CSV', 'No lines yet. Pick a product in column A and type the Quantity.', ui.ButtonSet.OK);
  } else {
    ui.alert('Purchase CSV', '✔ ' + res.rows.length + ' line(s) are ready to download.', ui.ButtonSet.OK);
  }
}


// (no trailing _ : called from the browser window) Builds BOTH files in memory (nothing is saved in Drive): the purchase file for the POS and the product xlsx of the
// products on the lines. Returns {problems:[…]} | {empty:true} | {files:[{name, mime, b64}], count, products}.
function pcBuildFiles() {
  const res = pcCollect_();
  if (res.problems.length) return { problems: res.problems };
  if (!res.rows.length) return { empty: true };
  return pcFilesFromRows_(res.rows);
}

// Both files from purchase rows [{code, cost (incl tax), qty, variant, tax, disc, exp}] (from the sheet or the window list).
function pcFilesFromRows_(rows) {
  const res = { rows: rows };
  const list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Product List');
  const n = Math.max(list.getLastRow() - 1, 1);
  const bn = list.getRange(2, 2, n, 2).getDisplayValues();
  const nm = list.getRange(2, 5, n, 2).getValues();
  const prodAoa = pcProductAoa_(res.rows, { bar: bn.map(r => r[0]), name: bn.map(r => r[1]), npp: nm.map(r => r[0]), mrp: nm.map(r => r[1]) });
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
  const enc = bytes => Utilities.base64Encode(bytes);
  return {
    count: res.rows.length,
    products: prodAoa.length - 1,
    files: [
      // same format as Product Tools → Export Products to XLS (.xls Excel 97-2003, sheet "Worksheet"), which the POS
      // product import already reads; an .xlsx of the same rows came in with NPP / MRP = 0 (23 Oct, user)
      { name: 'Products_' + stamp + '.xls', mime: 'application/vnd.ms-excel', b64: enc(pcXlsBytes_(prodAoa)) },
      { name: 'Purchase_' + stamp + '.xlsx', mime: PC_XLSX_MIME_, b64: enc(pcXlsxBytes_(pcAoa_(res.rows))) }
    ]
  };
}

// Browser side, shared by the small download window and the Loose purchase window: saves every file straight away
// (one after the other, no link to click) and keeps links as a back-up if the browser blocks a download.
const PC_SAVE_JS_ = 'function pcSaveFiles(files, box){'
  + 'box.innerHTML="";'
  + 'files.forEach(function(f,i){'
  + ' var bin=atob(f.b64),u=new Uint8Array(bin.length);for(var k=0;k<bin.length;k++)u[k]=bin.charCodeAt(k);'
  + ' var url=URL.createObjectURL(new Blob([u],{type:f.mime}));'
  + ' var a=document.createElement("a");a.href=url;a.download=f.name;a.textContent=f.name;a.style.display="block";a.style.margin="3px 0";'
  + ' box.appendChild(a);'
  + ' setTimeout(function(){a.click();},i*700);'
  + '});}';

// DOWNLOAD XLSX FILE button / menu: small window that builds both files and saves them at once.
function pcExportXlsx() {
  if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PC_SHEET_)) pcEnsureLayout_();
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:13px;padding:12px;">'
    + '<p id="m" style="font-size:14px">Making the files…</p>'
    + '<div id="f" style="font-size:12px"></div>'
    + '<p id="h" style="color:#777;font-size:11px;display:none">Not downloaded by itself? Click the file names above (or allow multiple downloads in the address bar).</p>'
    + '<button onclick="google.script.host.close()" style="margin-top:8px;padding:5px 14px">Close</button>'
    + '<script>' + PC_SAVE_JS_
    + 'google.script.run.withSuccessHandler(function(r){'
    + ' var m=document.getElementById("m");'
    + ' if(r.problems){m.style.color="#c00";m.innerHTML="Nothing downloaded - please fix:<br><span style=\'font-size:12px\'>"+r.problems.slice(0,10).join("<br>")+"</span>";return;}'
    + ' if(r.empty){m.style.color="#c00";m.textContent="There are no lines to download.";return;}'
    + ' m.innerHTML="✔ "+r.count+" purchase line(s), "+r.products+" product(s) - downloading "+r.files.length+" files…";'
    + ' pcSaveFiles(r.files,document.getElementById("f"));document.getElementById("h").style.display="block";'
    + '}).withFailureHandler(function(e){var m=document.getElementById("m");m.style.color="#c00";m.textContent=e.message;}).pcBuildFiles();'
    + '<\/script></div>'
  ).setWidth(420).setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(html, 'Download');
}

// Clears the typed lines (columns A–H, rows 5 down); keeps the formulas, dropdowns and the default tax.
function pcClearRows() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PC_SHEET_);
  if (!sh) return;
  if (ui.alert('Clear all purchase lines?', 'This empties Product, Quantity, Variant, Tax, Discount, Expiry, Cost override and the NPP / MRP copies on "' + PC_SHEET_ + '" (rows ' + PC_FIRST_ + ' down). The formulas stay.',
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  pcClearAllLines_(sh);
}

// Empties every typed purchase cell (formulas, dropdowns and the default tax stay). J / K are only copies: Product List keeps its NPP / MRP.
function pcClearAllLines_(sh) {
  [1, 4, 5, 6, 7, 8, PC_COST_COL_, PC_NPP_COL_, PC_MRP_COL_].forEach(c => sh.getRange(PC_FIRST_, c, PC_ROWS_, 1).clearContent());
  SpreadsheetApp.flush();
}
