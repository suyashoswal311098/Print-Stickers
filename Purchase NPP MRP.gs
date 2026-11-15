/***** PURCHASE CSV — NPP / MRP columns (J, K, after the Margin column I) that you can change in the sheet *****
 * J = NPP, K = MRP of the product on the line, copied from Product List (E, F) when the product is picked in A
 * (by hand, by paste, or by the purchase window). Typing a new NPP / MRP in J / K saves it straight into
 * Product List (first row with that barcode) and refreshes J / K on every other line of the same product.
 * Net Unit Cost (C) uses the NPP when there is no Cost override (N), so a new NPP changes C by itself.
 * Called from onEdit (Space Insert.gs); simple trigger, so it may only touch this spreadsheet.
 */

const PC_NPP_COL_ = 10;   // J
const PC_MRP_COL_ = 11;   // K

// Product List price lookup: purchase line text ("barcode name", or plain barcode / name) → {row, npp, mrp} | null.
function pcPriceIndex_() {
  const list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Product List');
  const n = Math.max(list.getLastRow() - 1, 1);
  const bn = list.getRange(2, 2, n, 2).getDisplayValues();
  const ef = list.getRange(2, 5, n, 2).getValues();
  const sq = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toUpperCase();
  const byBar = {}, byName = {};
  bn.forEach((r, i) => {
    const v = { row: i + 2, npp: ef[i][0], mrp: ef[i][1] };
    const k = normBar_(r[0]), nm = sq(r[1]);
    if (k && !byBar[k]) byBar[k] = v;
    if (nm && !byName[nm]) byName[nm] = v;
  });
  return function (text) {
    const full = sq(text);
    if (!full) return null;
    const m = full.match(/^(\S+)\s/);   // same as the B formula: the first word is the barcode
    return (m && byBar[normBar_(m[1])]) || byBar[normBar_(full)] || byName[full] || null;
  };
}

// Fills J / K of sheet rows r1..r2 from Product List (blank when the line has no product / product not found).
function pcFillPrices_(sh, r1, r2, look) {
  if (r2 < r1) return;
  look = look || pcPriceIndex_();
  const a = sh.getRange(r1, 1, r2 - r1 + 1, 1).getValues();
  sh.getRange(r1, PC_NPP_COL_, a.length, 2).setValues(a.map(r => { const p = look(r[0]); return p ? [p.npp, p.mrp] : ['', '']; }));
}

// After NPP / MRP of some products changed in Product List: refresh J / K on every purchase line of those products.
// rowsChanged = Product List row numbers.
function pcRefreshPrices_(rowsChanged) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PC_SHEET_);
  if (!sh || !rowsChanged.length) return;
  const want = {};
  rowsChanged.forEach(r => { want[r] = true; });
  const look = pcPriceIndex_();
  const a = sh.getRange(PC_FIRST_, 1, PC_ROWS_, 1).getValues();
  const rows = [];
  a.forEach((r, i) => { const p = String(r[0]).trim() && look(r[0]); if (p && want[p.row]) rows.push(PC_FIRST_ + i); });
  rows.forEach(row => pcFillPrices_(sh, row, row, look));
}

// onEdit part for the Purchase CSV sheet. Returns true when the edit was on that sheet (nothing else to do).
function pcOnEdit_(e) {
  const range = e && e.range;
  if (!range) return false;
  const sh = range.getSheet();
  if (sh.getName() !== PC_SHEET_) return false;
  const r1 = Math.max(range.getRow(), PC_FIRST_), r2 = Math.min(range.getLastRow(), PC_FIRST_ + PC_ROWS_ - 1);
  const c1 = range.getColumn(), c2 = range.getLastColumn();
  if (r1 > r2) return true;
  const rows = [];
  for (let r = r1; r <= r2; r++) rows.push(r);

  // something typed in a formula column (B, I Margin, L–Q info; C only when a pasted row starts in A -
  // a cost typed in C alone is handled below): emptied again so the
  // header-row formula keeps filling it
  const spill = PC_SPILL_COLS_.filter(c => (c !== 3 || c1 <= 1) && c >= c1 && c <= c2);
  spill.forEach(c => sh.getRange(r1, c, r2 - r1 + 1, 1).clearContent());
  if (spill.length && c1 === c2) { SpreadsheetApp.getActiveSpreadsheet().toast('This column is filled by a formula - type in the white / coloured ✎ columns.', 'Purchase CSV', 4); return true; }

  // product picked / changed / cleared in A → copy its NPP / MRP into J / K
  if (c1 <= 1) { pcFillPrices_(sh, r1, r2); return true; }

  // Net Unit Cost typed in C (user, 24 Oct): C is the header-row formula, so the typed number goes to the Cost override
  // (N, incl tax) and C is emptied again so the formula shows it. Blank (Delete) = nothing changes.
  if (c1 <= 3 && c2 >= 3) {
    const typed = sh.getRange(r1, 3, r2 - r1 + 1, 1).getValues();
    const bad = [];
    let saved = 0;
    typed.forEach((t, i) => {
      const row = r1 + i;
      if (t[0] === '' || t[0] === null) return;
      const v = pcNum_(t[0]);
      if (!String(sh.getRange(row, 1).getValue()).trim()) bad.push('Row ' + row + ': pick a product first');
      else if (isNaN(v) || v < 0) bad.push('Row ' + row + ': Net Unit Cost must be a number (0 or more)');
      else { sh.getRange(row, PC_COST_COL_).setValue(v); saved++; }
    });
    sh.getRange(r1, 3, r2 - r1 + 1, 1).clearContent();   // the formula in C4 fills the column again
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (bad.length) ss.toast(bad.slice(0, 3).join('\n'), 'Net Unit Cost', 6);
    else if (saved) ss.toast('Saved as Cost override (column N). Clear N to use NPP again.', 'Net Unit Cost', 4);
    return true;
  }

  // NPP / MRP typed in J / K → save in Product List
  if (c2 < PC_NPP_COL_ || c1 > PC_MRP_COL_) return true;
  const look = pcPriceIndex_();
  const list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Product List');
  const changed = [], bad = [];
  rows.forEach(row => {
    const p = look(sh.getRange(row, 1).getValue());
    for (let c = Math.max(c1, PC_NPP_COL_); c <= Math.min(c2, PC_MRP_COL_); c++) {
      const cell = sh.getRange(row, c), v = pcNum_(cell.getValue());
      const what = c === PC_NPP_COL_ ? 'NPP' : 'MRP';
      if (!p) { cell.clearContent(); if (String(sh.getRange(row, 1).getValue()).trim()) bad.push('Row ' + row + ': product not in Product List'); continue; }
      if (isNaN(v) || v < 0) {   // blank / text / negative → put the Product List value back
        cell.setValue(c === PC_NPP_COL_ ? p.npp : p.mrp);
        bad.push('Row ' + row + ': ' + what + ' must be a number (0 or more) - not saved');
        continue;
      }
      list.getRange(p.row, c === PC_NPP_COL_ ? 5 : 6).setValue(v);
      changed.push(p.row);
    }
  });
  if (changed.length) pcRefreshPrices_(changed);   // same product on other lines
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (bad.length) ss.toast(bad.slice(0, 3).join('\n'), 'NPP / MRP', 6);
  else if (changed.length) ss.toast('Saved in Product List.', 'NPP / MRP', 3);
  return true;
}
