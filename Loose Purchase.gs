/***** PURCHASE WINDOW — LOOSE (buy by weight, split into the Product List's pack sizes) or PACKED (scan barcode /
 * type name, packs + rate per pack); both put the lines in "Purchase CSV" *****
 * Product Tools → Purchase CSV for POS → "Add loose purchase…" opens "Loose Purchase Dialog".
 * The window does the thinking (find the pack sizes in the product names, split the kg, tax calculator).
 * This file only (1) hands the window the Product List and (2) writes the finished lines into the sheet:
 *   A = dropdown text "barcode name", D = packs, F = tax code, N = Cost override (cost per pack INCLUDING tax; the window
 *   works ex tax and the server adds the tax with pcInclCost_).
 * B / C / I / L / M / O / P / Q are the header-row formulas of the sheet and are never written here.
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
    defTax: String(sh.getRange('B2').getValue() || '').trim(),
    draft: PropertiesService.getUserProperties().getProperty(lpDraftKey_()) || ''
  };
}

// Window lines NOT yet in the sheet, kept so they come back after the window / page is closed (user, 25 Oct).
// Per user + spreadsheet; json = '{"ts":…,"lines":[…]}' or '' to forget. Max ~9 KB per property (≈ 60 lines).
function lpDraftKey_() { return 'lp_draft_' + SpreadsheetApp.getActiveSpreadsheet().getId(); }
function lpSaveDraft(json) {
  const props = PropertiesService.getUserProperties();
  if (!json) props.deleteProperty(lpDraftKey_());
  else props.setProperty(lpDraftKey_(), json);
  return true;
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
  sh.getRange(first, PC_COST_COL_, n, 1).setValues(lines.map(l => [pcInclCost_(l.cost, l.tax)]));
  // H = Expiry (text MM/dd/yyyy, today + shelf life, worked out in the window); blank when no shelf life
  sh.getRange(first, 8, n, 1).setNumberFormat('@').setValues(lines.map(l => [/^\d{2}\/\d{2}\/\d{4}$/.test(String(l.exp || '')) ? l.exp : '']));   // N = cost per pack INCLUDING tax (the window works ex tax)
  pcFillPrices_(sh, first, last);   // J / K = NPP / MRP (a script write does not run onEdit)
  SpreadsheetApp.flush();
  return { first: first, last: last };
}

// "Download both files" in the window (user, 25 Oct: straight from the window list, no need to put it in the sheet first).
// lines = the window list [{dd, packs, cost (ex tax), tax, exp}]; none → the sheet's lines (as the DOWNLOAD XLSX FILE button).
// prices = NPP / MRP changes typed in the window and not saved yet: saved FIRST, so the Products file carries them.
function lpDownload(lines, prices) {
  if (prices && prices.length) { lpSavePrices(prices); SpreadsheetApp.flush(); }
  if (!lines || !lines.length) return pcBuildFiles();
  const look = pcIndex_(), taxOk = pcTaxCodes_(), problems = [], rows = [];
  lines.forEach((l, i) => {
    const p = look(l.dd), bad = [];
    if (!p) bad.push('product not found in Product List');
    if (!(l.packs > 0)) bad.push('packs must be above 0');
    if (!(l.cost >= 0)) bad.push('no cost');
    if (l.tax && taxOk.indexOf(l.tax) < 0) bad.push('tax code "' + l.tax + '" is not a POS code');
    if (bad.length) { problems.push('Line ' + (i + 1) + ' (' + l.dd + '): ' + bad.join('; ')); return; }
    rows.push({ code: p.code, cost: pcInclCost_(l.cost, l.tax), qty: l.packs, variant: '', tax: l.tax || '', disc: '', exp: /^\d{2}\/\d{2}\/\d{4}$/.test(String(l.exp || '')) ? l.exp : '' });
  });
  if (problems.length) return { problems: problems };
  return pcFilesFromRows_(rows);
}

// ✎ in the window, or the same product added again: change packs / cost / tax (/ expiry) of a line already in the sheet (row must still hold the same product).
function lpUpdate(row, dd, v) {
  const sh = lpRowCheck_(row, dd);
  if (!(v.packs > 0) || !(v.cost >= 0)) throw new Error('Packs and cost are not valid.');
  if (v.tax && pcTaxCodes_().indexOf(v.tax) < 0) throw new Error('Tax code "' + v.tax + '" is not a POS code.');
  sh.getRange(row, 4).setValue(v.packs);
  sh.getRange(row, 6).setValue(v.tax || '');
  sh.getRange(row, PC_COST_COL_).setValue(pcInclCost_(v.cost, v.tax));
  if (v.exp !== undefined) sh.getRange(row, 8).setNumberFormat('@').setValue(/^\d{2}\/\d{2}\/\d{4}$/.test(String(v.exp)) ? v.exp : '');   // same product added again: new expiry
  SpreadsheetApp.flush();
  return true;
}

// Lines already in the sheet, for the window's list (user, 24 Oct: change / remove them in the window too).
// [{row, dd, packs, cost (Net Unit Cost C, INCL tax), tax (F, or the default B2), exp (H)}], rows with a product in A only.
function lpReadSheet() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PC_SHEET_);
  if (!sh) return [];
  SpreadsheetApp.flush();
  const v = sh.getRange(PC_FIRST_, 1, PC_ROWS_, 8).getValues(), h = sh.getRange(PC_FIRST_, 8, PC_ROWS_, 1).getDisplayValues();
  const defTax = String(sh.getRange('B2').getValue() || '').trim();
  const out = [];
  v.forEach((r, i) => {
    const dd = String(r[0]).trim();
    if (!dd) return;
    out.push({ row: PC_FIRST_ + i, dd: dd, packs: pcNum_(r[3]), cost: pcNum_(r[2]), tax: String(r[5]).trim() || defTax, exp: String(h[i][0]).trim() });
  });
  return out;
}

// ✕ on a line already in the sheet: removes it and moves the lines below up one row (no gap). Only the typed
// columns move (A, D–H, J–K, N); the header-row formulas follow by themselves. The window reloads its list after this.
function lpClear(row, dd) {
  const sh = lpRowCheck_(row, dd);
  const end = PC_FIRST_ + PC_ROWS_ - 1;
  [[1, 1], [4, 5], [PC_NPP_COL_, 2], [PC_COST_COL_, 1]].forEach(([c, w]) => {
    if (row < end) {
      const below = sh.getRange(row + 1, c, end - row, w);
      sh.getRange(row, c, end - row, w).setValues(below.getValues());
    }
    sh.getRange(end, c, 1, w).clearContent();
  });
  SpreadsheetApp.flush();
  return true;
}

// "🗑 Clear all lines" in the window (it asks "sure?" first): empties all purchase lines in the sheet.
function lpClearAll() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PC_SHEET_);
  if (sh) pcClearAllLines_(sh);
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
  const rows = [];
  (changes || []).forEach(c => {
    const row = at[normBar_(c.code)];
    if (!row) return;
    const ok = v => typeof v === 'number' && isFinite(v) && v >= 0;
    if (ok(c.npp)) list.getRange(row, 5).setValue(c.npp);
    if (ok(c.mrp)) list.getRange(row, 6).setValue(c.mrp);
    if (ok(c.npp) || ok(c.mrp)) { done++; rows.push(row); }
  });
  pcRefreshPrices_(rows);   // NPP / MRP columns of the Purchase CSV lines
  return done;
}
