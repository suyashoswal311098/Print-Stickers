/***** ADD MANY — pick many products + qty in one window, then insert them all *****
 * Button "ADD MANY" on Print Console (left of LOAD) and Product Tools menu → openAddMany.
 * Each ticked product goes through the SAME insert as the INSERT button (insert3, silent),
 * Fast path writes all rows at once (checked against the sheet's own A2/H2 formulas first);
 * slow path = B2 + C2 filled, formulas recalculated, insert3 run, one by one. Packing size = 1.
 */

const AM_BUTTON_TITLE_ = 'ADD MANY';
const AM_BUTTON_PNG_ = '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';

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
    // Product List B (barcode), C (name), H (colour) — first row per name
    const n = Math.max(list.getLastRow() - 1, 1);
    const pl = list.getRange(2, 1, n, 9).getValues();
    const byName = {}, byBar = {};
    pl.forEach(r => {
      const k = normName_(r[2]), bk = normBar_(r[1]);
      const v = { bar: r[1], colour: r[7], name: normName_(r[2]) };
      if (k && !byName[k]) byName[k] = v;
      if (bk && !byBar[bk]) byBar[bk] = v;
    });
    // Dropdown names may be "<barcode> <name>" (e.g. "49876602 PUMPKIN SEEDS 100GM")
    const lookup = nm => {
      const full = normName_(nm);
      if (byName[full]) return byName[full];
      const m = full.match(/^(\S+)\s+(.+)$/);
      const p = m && byBar[normBar_(m[1])];
      return p && (!p.name || p.name === m[2]) ? p : null;
    };
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

// Puts the ADD MANY button on Print Console once (left of LOAD). Safe to run again.
function placeAddManyButton(quiet) {
  const work = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Print Console');
  if (!work) return;
  const has = work.getImages().some(img => img.getAltTextTitle() === AM_BUTTON_TITLE_);
  if (has) {
    if (!quiet) SpreadsheetApp.getActiveSpreadsheet().toast('The ADD MANY button is already on Print Console.', 'Add many', 5);
    return;
  }
  const blob = Utilities.newBlob(Utilities.base64Decode(AM_BUTTON_PNG_), 'image/png', 'add-many.png');
  const img = work.insertImage(blob, 8, 3, 3, 2); // column H, row 3
  img.setWidth(82).setHeight(38);
  img.setAltTextTitle(AM_BUTTON_TITLE_);
  img.assignScript('openAddMany');
  if (!quiet) SpreadsheetApp.getActiveSpreadsheet().toast('ADD MANY button placed. You can drag it anywhere.', 'Add many', 5);
}
