/***** PURCHASE CSV — build the POS "Add Purchase by CSV" file from the Product List *****
 * Sheet "Purchase CSV": pick a product in column A (dropdown = the Print Console B2 list, "barcode name",
 * type to search). FORMULAS IN THE HEADER ROW (no script on edit) fill Product Code (the barcode in front
 * of the picked text, so leading zeros stay) + Net Unit Cost (NPP from Product List, or the Cost override
 * in column K). You type Quantity, optionally Variant / Tax Rate Code / Discount / Expiry.
 * Tax Rate Code left empty = the Default Tax Rate Code in B2 (applied in the file).
 * Menu Product Tools → Purchase CSV for POS → Download XLSX gives a file with exactly the POS columns:
 *   Product Code | Net Unit Cost | Quantity | Product Variant | Tax Rate Code | Discount | Expiry
 * Sheet columns: A Product | B..H = the 7 POS columns | I Total ex tax (info) | J Check (info) | K Cost override (optional) | L Tax used (info) | M Rate incl tax (info) | N Total incl tax (info)
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
// true = the file's Net Unit Cost is the rate INCLUDING tax (user's choice, 22 Oct); false = the ex-tax cost from column C.
const PC_FILE_COST_INCL_ = true;
const PC_EXPIRY_FORMAT_ = 'MM/dd/yyyy';   // as in the POS sample file (09/26/2019); a date typed in Expiry is written like this
const PC_XLSX_MIME_ = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// Cost × (1 + tax %) for a Tax Rate Code (POS Tax Codes table); no / unknown code = no tax. 4 decimals.
function pcInclCost_(cost, taxCode) {
  const t = PC_TAX_CODES_.filter(x => x[1] === taxCode)[0];
  return Math.round(cost * (1 + (t ? t[2] : 0) / 100) * 10000) / 10000;
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
  if (sh.getMaxColumns() < 14) sh.insertColumnsAfter(sh.getMaxColumns(), 14 - sh.getMaxColumns());

  sh.getRange('A1').setValue('PURCHASE FILE FOR POS  (Add Purchase by CSV)').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setValue('Default Tax Rate Code  →').setHorizontalAlignment('right').setFontWeight('bold');
  sh.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(pcTaxCodes_(), true).setAllowInvalid(false).build()).setBackground('#fff2cc');
  sh.getRange('G2').setValue('Total ex tax').setHorizontalAlignment('right').setFontWeight('bold');
  sh.getRange('H2').setFormula('=SUM(I' + PC_FIRST_ + ':I' + lastRow + ')').setNumberFormat('#,##0.00').setFontWeight('bold');
  sh.getRange('M2').setValue('Total incl tax').setHorizontalAlignment('right').setFontWeight('bold');
  sh.getRange('N2').setFormula('=SUM(N' + PC_FIRST_ + ':N' + lastRow + ')').setNumberFormat('#,##0.00').setFontWeight('bold');
  sh.getRange('A3').setValue('Pick/type a product in column A and type Quantity. B, C, I, J, L, M, N are formulas in the header row (do not type there). ' +
    'B–H go into the file in the POS order (the file\'s Net Unit Cost = Rate incl tax, column M); the grey columns are only for you.').setFontColor('#666');

  // Header row: plain labels, except B, C, I, J = label + ONE array formula that fills the whole column.
  // Spill cells below must be empty, so old values/formulas in those columns are cleared first.
  [2, 3, 9, 10, 12, 13, 14].forEach(c => sh.getRange(PC_FIRST_, c, PC_ROWS_, 1).clearContent());
  const A = 'A' + PC_FIRST_ + ':A' + lastRow, B = 'B' + PC_FIRST_ + ':B' + lastRow, C = 'C' + PC_FIRST_ + ':C' + lastRow;
  const D = 'D' + PC_FIRST_ + ':D' + lastRow, K = 'K' + PC_FIRST_ + ':K' + lastRow;
  const bars = "'Product List'!$B$2:$B&\"\"";
  sh.getRange(PC_HEAD_ROW_, 1, 1, 14).setNumberFormat('General');
  sh.getRange(PC_HEAD_ROW_, 1, 1, 14).setValues([['Product (barcode + name)'].concat(PC_HEADERS_, ['', '', 'Cost override (optional)', '', '', ''])]);
  sh.getRange(PC_HEAD_ROW_, 2).setFormula('={"Product Code";ARRAYFORMULA(IF(' + A + '="","",IFERROR(REGEXEXTRACT(' + A + '&"","^\\S+"),"")))}');
  sh.getRange(PC_HEAD_ROW_, 3).setFormula('={"Net Unit Cost";ARRAYFORMULA(IF(' + B + '="","",IF(' + K + '<>"",' + K + ',IFERROR(VLOOKUP(' + B +
    ',{' + bars + ",'Product List'!$E$2:$E},2,FALSE),\"\"))))}");
  sh.getRange(PC_HEAD_ROW_, 9).setFormula('={"Total ex tax (info)";ARRAYFORMULA(IF(ISNUMBER(' + C + ')*ISNUMBER(' + D + '),' + C + '*' + D + ',""))}');
  sh.getRange(PC_HEAD_ROW_, 10).setFormula('={"Check (info)";ARRAYFORMULA(IF(' + A + '="","",IF(ISNA(VLOOKUP(' + B + ',{' + bars + ',' + bars + '},2,FALSE)),"⚠ not in Product List",' +
    'IF(NOT(ISNUMBER(' + D + ')),"⚠ type Quantity",IF(' + D + '<=0,"⚠ Quantity must be above 0",IF(NOT(ISNUMBER(' + C + ')),"⚠ no cost - type a Cost override","✔ ok"))))))}');
  sh.getRange(PC_HEAD_ROW_, 1, 1, 14).setBackground('#4a86c8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(PC_HEAD_ROW_, 1).setBackground('#1f3d66');
  sh.getRange(PC_HEAD_ROW_, 12).setFormula('={"Tax used in file (info)";ARRAYFORMULA(IF(' + A + '="","",IF(F' + PC_FIRST_ + ':F' + lastRow + '<>"",F' + PC_FIRST_ + ':F' + lastRow + ',IF($B$2="","⚠ no tax",$B$2))))}');
  const L = 'L' + PC_FIRST_ + ':L' + lastRow, M = 'M' + PC_FIRST_ + ':M' + lastRow;
  const taxPct = 'VLOOKUP(' + L + ",'" + PC_TAX_SHEET_ + "'!$B$2:$C$" + (PC_TAX_CODES_.length + 1) + ',2,FALSE)';
  sh.getRange(PC_HEAD_ROW_, 13).setFormula('={"Rate incl tax (info)";ARRAYFORMULA(IF(ISNUMBER(' + C + '),IFERROR(ROUND(' + C + '*(1+' + taxPct + '/100),2),""),""))}');
  sh.getRange(PC_HEAD_ROW_, 14).setFormula('={"Total incl tax (info)";ARRAYFORMULA(IF(ISNUMBER(' + C + ')*ISNUMBER(' + D + '),IFERROR(ROUND(' + C + '*' + D + '*(1+' + taxPct + '/100),2),""),""))}');
  sh.getRange(PC_HEAD_ROW_, 9, 1, 6).setBackground('#7f7f7f');
  sh.setFrozenRows(PC_HEAD_ROW_);

  // formats: variant + discount + expiry as plain text ("5%" stays as typed); code/cost come from the formulas
  sh.getRange(PC_FIRST_, 2, PC_ROWS_, 1).setNumberFormat('General');
  sh.getRange(PC_FIRST_, 5, PC_ROWS_, 1).setNumberFormat('@');
  sh.getRange(PC_FIRST_, 7, PC_ROWS_, 1).setNumberFormat('@');
  sh.getRange(PC_FIRST_, 3, PC_ROWS_, 1).setNumberFormat('0.00');
  sh.getRange(PC_FIRST_, 4, PC_ROWS_, 1).setNumberFormat('0.##');
  sh.getRange(PC_FIRST_, 8, PC_ROWS_, 1).setNumberFormat('@');
  sh.getRange(PC_FIRST_, 9, PC_ROWS_, 1).setNumberFormat('#,##0.00');
  sh.getRange(PC_FIRST_, 11, PC_ROWS_, 1).setNumberFormat('0.00');

  // dropdowns
  sh.getRange(PC_FIRST_, 1, PC_ROWS_, 1).setDataValidation(rule);
  sh.getRange(PC_FIRST_, 6, PC_ROWS_, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(pcTaxCodes_(), true).setAllowInvalid(false).build());

  sh.getRange(PC_FIRST_, 9, PC_ROWS_, 6).setFontColor('#666');
  sh.getRange(PC_FIRST_, 13, PC_ROWS_, 2).setNumberFormat('#,##0.00');
  sh.getRange(PC_FIRST_, 2, PC_ROWS_, 7).setHorizontalAlignment('left');

  const jr = sh.getRange(PC_FIRST_, 10, PC_ROWS_, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('⚠').setFontColor('#cc0000').setBold(true).setRanges([jr]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('✔').setFontColor('#1e8e3e').setRanges([jr]).build()
  ]);

  sh.setColumnWidth(1, 380);
  [110, 95, 75, 100, 110, 80, 95, 95, 230, 175, 170, 150, 160].forEach((w, i) => sh.setColumnWidth(i + 2, w));
  sh.activate();
  sh.getRange(PC_FIRST_, 1).activate();
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
    if (isNaN(cost) || cost < 0) bad.push('Net Unit Cost is empty (no NPP in Product List - type a Cost override in column K)');
    if (tax && taxOk.indexOf(tax) < 0) bad.push('Tax Rate Code "' + tax + '" is not a POS code');
    if (bad.length) { problems.push('Row ' + row + ': ' + bad.join('; ')); return; }
    const exp = r[7] instanceof Date ? Utilities.formatDate(r[7], tz, PC_EXPIRY_FORMAT_) : String(r[7]).trim();
    rows.push({ row: row, code: c, cost: PC_FILE_COST_INCL_ ? pcInclCost_(cost, tax) : cost, qty: qty, variant: String(r[4]).trim(), tax: tax, disc: String(r[6]).trim(), exp: exp });
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
function pcXlsxBytes_(aoa) {
  const lib = sheetJs_();
  const ws = lib.utils.aoa_to_sheet(aoa);
  for (let r = 2; r <= aoa.length; r++) {
    ['A', 'D', 'E', 'G'].forEach(col => {
      const c = ws[col + r];
      if (c) { c.t = 's'; c.v = String(c.v); c.z = '@'; delete c.w; }
    });
  }
  const wb = lib.utils.book_new();
  lib.utils.book_append_sheet(wb, ws, 'sample_purchase_products');   // same sheet name as the POS sample
  const u8 = new Uint8Array(lib.write(wb, { bookType: 'xlsx', type: 'array', bookSST: true }));
  return Array.prototype.map.call(u8, b => b > 127 ? b - 256 : b);
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

// The POS page accepts .xlsx only ("Invalid file selected. Only .xlsx file is allowed"), so there is no CSV download.
function pcExportXlsx() { pcExport_('xlsx'); }

// Builds the .xlsx in Drive (temp file) from the sheet. Returns {problems:[…]} | {empty:true} | {name,url,fileId,count}.
function pcMakeFile_() {
  const res = pcCollect_();
  if (res.problems.length) return { problems: res.problems };
  if (!res.rows.length) return { empty: true };
  const aoa = pcAoa_(res.rows);
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
  const name = 'Purchase_' + stamp + '.xlsx';
  const file = DriveApp.createFile(Utilities.newBlob(pcXlsxBytes_(aoa), PC_XLSX_MIME_, name));
  const fileId = file.getId();
  scheduleExportCleanup_(fileId);   // same temp-file clean-up as Export Products to XLS
  return { name: name, url: file.getDownloadUrl(), fileId: fileId, count: res.rows.length };
}

function pcExport_() {
  const ui = SpreadsheetApp.getUi();
  const r = pcMakeFile_();
  if (r.problems) {
    ui.alert('Nothing downloaded - please fix', r.problems.slice(0, 15).join('\n') +
      (r.problems.length > 15 ? '\n…and ' + (r.problems.length - 15) + ' more' : ''), ui.ButtonSet.OK);
    return;
  }
  if (r.empty) { ui.alert('There are no lines to download.'); return; }
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:13px;padding:10px;">'
    + '<p><b>' + r.count + '</b> purchase line(s) ready for the POS.</p>'
    + '<p><a id="dl" href="' + r.url + '" target="_blank">Download ' + r.name + '</a></p>'
    + '<p style="color:#555;font-size:12px;">POS → Purchases → Add Purchase by CSV → Browse → pick this file.</p>'
    + '<p id="msg" style="color:#777;font-size:11px;">The file is removed from Drive 15 seconds after you click the link.</p>'
    + '<script>'
    + 'document.getElementById("dl").addEventListener("click", function(){'
    + '  var n = 15; var m = document.getElementById("msg");'
    + '  var t = setInterval(function(){'
    + '    n--; m.textContent = "Removing from Drive in " + n + "s...";'
    + '    if (n <= 0) { clearInterval(t);'
    + '      google.script.run'
    + '        .withSuccessHandler(function(){ m.textContent = "Removed from Drive."; })'
    + '        .withFailureHandler(function(e){ m.textContent = "Could not remove: " + e.message; })'
    + '        .deleteExportFile("' + r.fileId + '"); }'
    + '  }, 1000);'
    + '});'
    + '<\/script></div>'
  ).setWidth(400).setHeight(230);
  ui.showModalDialog(html, 'Purchase file ready');
}

// Clears the typed lines (columns A–H, rows 5 down); keeps the formulas, dropdowns and the default tax.
function pcClearRows() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PC_SHEET_);
  if (!sh) return;
  if (ui.alert('Clear all purchase lines?', 'This empties Product, Quantity, Variant, Tax, Discount, Expiry and Cost override on "' + PC_SHEET_ + '" (rows ' + PC_FIRST_ + ' down). The formulas stay.',
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  [1, 4, 5, 6, 7, 8, 11].forEach(c => sh.getRange(PC_FIRST_, c, PC_ROWS_, 1).clearContent());
}
