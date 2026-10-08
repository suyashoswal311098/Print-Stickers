/***** ADD MANY — pick many products + qty in one window, then insert them all *****
 * Button "ADD MANY" on Print Console (left of LOAD) and Product Tools menu → openAddMany.
 * Each ticked product goes through the SAME insert as the INSERT button (insert3, silent),
 * Fast path writes all rows at once (checked against the sheet's own A2/H2 formulas first);
 * slow path = B2 + C2 filled, formulas recalculated, insert3 run, one by one. Packing size = 1.
 */

const AM_BUTTON_TITLE_ = 'ADD MANY';
const AM_BUTTON_PNG_ = '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'; // 328x152 (4x), shown at 82x38

// Product List index: lookup(dropdown name) → {bar, colour, name, row}. Dropdown names may be
// the plain name or "<barcode> <name>" (e.g. "49876602 PUMPKIN SEEDS 100GM").
function amProductIndex_(list) {
  const n = Math.max(list.getLastRow() - 1, 1);
  const pl = list.getRange(2, 1, n, 9).getValues();
  const byName = {}, byBar = {};
  pl.forEach((r, i) => {
    const k = normName_(r[2]), bk = normBar_(r[1]);
    const v = { bar: r[1], colour: r[7], name: k, row: i + 2 };
    if (k && !byName[k]) byName[k] = v;
    if (bk && !byBar[bk]) byBar[bk] = v;
  });
  return function lookup(nm) {
    const full = normName_(nm);
    if (byName[full]) return byName[full];
    const m = full.match(/^(\S+)\s+(.+)$/);
    const p = m && byBar[normBar_(m[1])];
    return p && (!p.name || p.name === m[2]) ? p : null;
  };
}

function openAddMany() {
  const html = HtmlService.createHtmlOutputFromFile('Add Many Dialog')
    .setWidth(560)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add many stickers');
}

// Product names = the B2 dropdown's own list (same as pullFromInventory), plus stickers already added.
function amGetProducts() {
  const work = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Print Console');
  const rule = work.getRange('B2').getDataValidation();
  let names = [];
  if (rule) {
    const type = rule.getCriteriaType();
    const vals = rule.getCriteriaValues();
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      names = vals[0].getValues().flat();
    } else if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      names = vals[0];
    }
  }
  const seen = {};
  names = names.map(v => String(v).trim()).filter(v => {
    if (!v || seen[v]) return false;
    seen[v] = true;
    return true;
  });
  // Show products in Product List order (the order saved by Submit), unknown ones last
  try {
    const lookup = amProductIndex_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Product List'));
    const pos = {};
    names.forEach((nm, i) => { const p = lookup(nm); pos[nm] = p ? p.row : 1e9 + i; });
    names.sort((a, b) => pos[a] - pos[b]);
  } catch (err) { console.warn('Add many: sort by Product List failed: ' + err.message); }
  return { names: names, used: Number(work.getRange('C5').getValue()) || 0, limit: 240 };
}

// items = [{name, qty}] → inserts each like the INSERT button. Returns a summary.
// FAST: barcode + colour are looked up in Product List (by name, or by "<barcode> <name>") and all
// rows are written in one go. Safety check first: the first product is put in B2 and the sheet's own
// A2 (barcode) / H2 (colour) formulas must give the same answer; if not, the slow per-item way is used.
function amInsertMany(items) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const work = ss.getSheetByName('Print Console');
  const list = ss.getSheetByName('Product List');
  items = (items || []).filter(it => it && String(it.name).trim() && Number(it.qty) > 0);
  if (!items.length) return { inserted: 0, skipped: [] };

  const used = Number(work.getRange('C5').getValue()) || 0;
  const total = items.reduce((s, it) => s + Number(it.qty), 0);
  if (used + total > 240) {
    return { inserted: 0, skipped: [], error: 'Too many stickers: ' + used + ' already added + ' + total +
      ' new = ' + (used + total) + '. The limit is 240.' };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { inserted: 0, skipped: [], error: 'Another insert is running. Please try again.' };
  let fastOk = false, plan = [], notFound = [];
  const skipped = [];
  try {
    const lookup = amProductIndex_(list);
    items.forEach(it => {
      const p = lookup(it.name);
      if (p && p.bar !== '' && p.bar !== null) plan.push({ name: it.name, qty: Number(it.qty), bar: p.bar, colour: p.colour });
      else notFound.push({ name: it.name, qty: Number(it.qty) }); // left to the slow way below
    });

    if (plan.length) {
      // Safety check with the sheet's own formulas
      work.getRange('B2:C2').setValues([[String(plan[0].name).trim(), plan[0].qty]]);
      SpreadsheetApp.flush();
      const top = work.getRange('A2:H2').getValues()[0];
      fastOk = normBar_(top[0]) === normBar_(plan[0].bar) && String(top[7]) === String(plan[0].colour);
      if (!fastOk) console.warn('Add many: fast check failed', top[0], plan[0].bar, top[7], plan[0].colour);
      work.getRange('B2:C2').clearContent();
      work.getRange('D2').setValue(1);
    }

    if (fastOk) {
      // Same row as insert3 builds: A barcode, C qty, D packing 1, J colour (value + background), rest blank
      const row = Math.max(work.getLastRow() + 1, 8);
      const vals = plan.map(p => [p.bar, '', p.qty, 1, '', '', '', '', '', p.colour]);
      const bgs = plan.map(p => ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', p.colour || '#ffffff']);
      const rng = work.getRange(row, 1, plan.length, 10);
      rng.setValues(vals);
      rng.setBackgrounds(bgs);
    }
  } finally {
    lock.releaseLock();
  }

  let inserted = fastOk ? plan.length : 0;
  const slow = fastOk ? notFound : plan.concat(notFound);
  if (slow.length) {
    // Slow, sure way: exactly what the INSERT button does, one by one
    const cachedProdList = list.getRange(2, 2, Math.max(list.getLastRow() - 1, 1), 9).getValues();
    slow.forEach(it => {
      work.getRange('B2:D2').setValues([[String(it.name).trim(), it.qty, 1]]);
      SpreadsheetApp.flush(); // let A2 / H2 formulas follow B2
      let status;
      try {
        status = insert3(null, true, cachedProdList);
      } catch (err) {
        status = 'error: ' + err.message;
      }
      if (status === 'success') inserted++;
      else skipped.push(it.name + ' (' + (status === 'skipped' ? 'not found in Product List' : status) + ')');
    });
    work.getRange('B2:C2').clearContent();
    work.getRange('D2').setValue(1);
  }

  work.getRange('B2').activate();
  ss.toast('Inserted: ' + inserted + (skipped.length ? ' | Skipped: ' + skipped.length : ''), '✅ Add many', 6);
  return { inserted: inserted, skipped: skipped, fast: fastOk };
}

// Saves the Submit order in Product List: the products' rows are moved together, in this order,
// to where the first of them is now (whole rows, like dragging rows in the sheet).
function amSaveOrder(names) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const list = ss.getSheetByName('Product List');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { moved: 0, error: 'Another job is running. Order not saved.' };
  try {
    const lookup = amProductIndex_(list);
    const cur = [], seen = {};
    (names || []).forEach(nm => {
      const p = lookup(nm);
      if (p && !seen[p.row]) { seen[p.row] = 1; cur.push(p.row); }
    });
    if (cur.length < 2) return { moved: 0 };
    const anchor = Math.min.apply(null, cur);
    let moved = 0;
    for (let k = 0; k < cur.length; k++) {
      const target = anchor + k, from = cur[k];
      if (from === target) continue;
      // from is always below target (rows above are already placed), so this is an upward move
      list.moveRows(list.getRange(from + ':' + from), target);
      moved++;
      for (let j = k + 1; j < cur.length; j++) if (cur[j] >= target && cur[j] < from) cur[j]++;
    }
    return { moved: moved, from: anchor, to: anchor + cur.length - 1 };
  } finally {
    lock.releaseLock();
  }
}

// Puts a picture button on Print Console. If it is already there (same alt title) it is
// swapped for the current picture in the same place and size (so moved buttons stay put).
function amPlaceButton_(title, png, col, row, offX, offY, fn) {
  const work = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Print Console');
  if (!work) return '';
  let w = 82, h = 38;
  const old = work.getImages().filter(img => img.getAltTextTitle() === title);
  if (old.length) {
    const o = old[0], a = o.getAnchorCell();
    col = a.getColumn(); row = a.getRow(); offX = o.getAnchorCellXOffset(); offY = o.getAnchorCellYOffset();
    w = o.getWidth(); h = o.getHeight();
    old.forEach(img => img.remove());
  }
  const blob = Utilities.newBlob(Utilities.base64Decode(png), 'image/png', title.toLowerCase().replace(/\s+/g, '-') + '.png');
  const img = work.insertImage(blob, col, row, offX, offY);
  img.setWidth(w).setHeight(h);
  img.setAltTextTitle(title);
  img.assignScript(fn);
  return old.length ? 'updated' : 'placed';
}

function placeAddManyButton(quiet) {
  const r = amPlaceButton_(AM_BUTTON_TITLE_, AM_BUTTON_PNG_, 8, 3, 3, 2, 'openAddMany'); // column H, row 3
  if (!quiet) SpreadsheetApp.getActiveSpreadsheet().toast('ADD MANY button ' + r + '. You can drag it anywhere.', 'Add many', 5);
}
