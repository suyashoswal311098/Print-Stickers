/**
 * UPDATE PRODUCT LIST — v5
 * Detection is done entirely in Apps Script (the Helper column on the sheet
 * is just a visual preview and is NOT read here).
 *
 * STEP 0 — Duplicate check within the Import List itself:
 *   - BARCODE duplicate: always flagged (a shared barcode is inherently
 *     ambiguous, regardless of name/price).
 *   - NAME duplicate: only flagged when NPP *and* MRP also match between the
 *     rows — same convention as the "Find Duplicates" cleanup tool below, so
 *     two different products that happen to share a generic name aren't
 *     wrongly flagged.
 *   Any row caught in a duplicate group is held back entirely (not matched,
 *   not appended) and shown in the sidebar, where you can delete the extra
 *   row(s) directly. Deletion re-verifies the row's content against a
 *   signature captured at scan time, so a stale sidebar can't delete the
 *   wrong row if the sheet changed underneath it. Fix these, then run again.
 *
 * Logic per remaining Import row (against Product List):
 *   1) Barcode matches an existing product
 *        - Name also matches  -> compare NPP & MRP, update only what differs (immediate).
 *        - Name differs        -> "Name Change" conflict -> held for sidebar review.
 *   2) Barcode is new but NAME matches       -> "Barcode Change" conflict -> held for sidebar review.
 *   3) Neither barcode nor name matches      -> genuinely new -> appended immediately.
 *
 * Column H (Colour Code) is never written, so its formula stays intact.
 *
 * Menu: see onOpen() at the bottom — adds "Product Tools" with entries for
 * both this import flow and the separate Product List "Find Duplicates"
 * cleanup tool (executeFindDuplicates), so the cleanup can be run any time
 * without needing to import anything first.
 */

function updateProductList2() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const ui           = SpreadsheetApp.getUi();
  const importSheet  = ss.getSheetByName('Import List');
  const productSheet = ss.getSheetByName('Product List');

  if (!importSheet || !productSheet) {
    ui.alert("Make sure both sheets exist.");
    return;
  }

  const HEADER_ROWS     = 3;
  const FIRST_ROW       = HEADER_ROWS + 1;
  const NUM_IMPORT_COLS = 9;
  const totalRows       = importSheet.getLastRow();
  const dataRows        = totalRows - HEADER_ROWS;
  if (dataRows <= 0) {
    ui.alert("No rows under header.");
    return;
  }
  importSheet.getRange(FIRST_ROW, 1, Math.max(importSheet.getMaxRows() - HEADER_ROWS, 1), 1)
             .setNumberFormat('@');
  productSheet.getRange(2, 2, Math.max(productSheet.getMaxRows() - 1, 1), 1)
             .setNumberFormat('@');

  const rawData = importSheet.getRange(FIRST_ROW, 1, dataRows, NUM_IMPORT_COLS).getValues();

  const rawBarDisp = importSheet.getRange(FIRST_ROW, 1, dataRows, 1).getDisplayValues();
  for (let i = 0; i < rawData.length; i++) {
    rawData[i][0] = rawBarDisp[i][0];
  }

  const lastProdRow = productSheet.getLastRow();
  const prodCols    = Math.max(productSheet.getLastColumn(), 9);
  const prodRows    = Math.max(0, lastProdRow - 1);
  const prodData    = prodRows ? productSheet.getRange(2, 1, prodRows, prodCols).getValues() : [];

  if (prodRows) {
    const prodBarDisp = productSheet.getRange(2, 2, prodRows, 1).getDisplayValues();
    for (let i = 0; i < prodData.length; i++) {
      prodData[i][1] = prodBarDisp[i][0];
    }
  }

  // Product List columns (0-based): B=1 bar, C=2 name, D=3 brand, E=4 npp, F=5 mrp, G=6 shelf, H=7 colour(protected), I=8 rating
  const COL_IN  = { bar:0, name:1, brand:2, npp:3, mrp:4, shelf:5, rating:6 };
  const COL_OUT = { bar:1, name:2, brand:3, npp:4, mrp:5, shelf:6, rating:8 };

  // --- STEP 0: Detect duplicates WITHIN the Import List itself ---
  const dupBarGroups  = new Map(); // normBar              -> [{rowNumber, itemCode, productName, brand, npp, mrp, sig}]
  const dupNameGroups = new Map(); // normName+NPP+MRP key -> [...]

  rawData.forEach((rowIn, i) => {
    const bar  = rowIn[COL_IN.bar];
    const name = rowIn[COL_IN.name];
    if (!bar && !name) return;

    const info = {
      rowNumber:   FIRST_ROW + i,
      itemCode:    bar,
      productName: name,
      brand:       rowIn[COL_IN.brand],
      npp:         rowIn[COL_IN.npp],
      mrp:         rowIn[COL_IN.mrp],
      sig:         importRowSig_(rowIn)
    };

    const bKey = normBar_(bar);
    const nKey = normName_(name);
    const nppN = normPrice_(rowIn[COL_IN.npp]);
    const mrpN = normPrice_(rowIn[COL_IN.mrp]);

    if (bKey) {
      if (!dupBarGroups.has(bKey)) dupBarGroups.set(bKey, []);
      dupBarGroups.get(bKey).push(info);
    }

    // Name duplicates only count when the price is known AND matches too.
    if (nKey && nppN !== "" && mrpN !== "") {
      const priceKey = nKey + "\u0001" + nppN + "\u0001" + mrpN;
      if (!dupNameGroups.has(priceKey)) dupNameGroups.set(priceKey, []);
      dupNameGroups.get(priceKey).push(info);
    }
  });

  const duplicateRowNumbers = new Set(); // rows to skip in the main matching loop below
  const duplicateConflicts  = [];        // sidebar cards for duplicates

  dupBarGroups.forEach((rows, key) => {
    if (rows.length > 1) {
      duplicateConflicts.push({ type: 'duplicate', subtype: 'barcode', key: key, rows: rows });
      rows.forEach(r => duplicateRowNumbers.add(r.rowNumber));
    }
  });
  dupNameGroups.forEach((rows) => {
    if (rows.length > 1) {
      duplicateConflicts.push({ type: 'duplicate', subtype: 'name', key: rows[0].productName, rows: rows });
      rows.forEach(r => duplicateRowNumbers.add(r.rowNumber));
    }
  });

  // Build lookup maps (normalized keys) against the existing Product List
  const barMap  = new Map(); // "barcode" -> row index in prodData
  const nameMap = new Map(); // "NAME"(upper) -> row index in prodData
  prodData.forEach((r, i) => {
    const b = normBar_(r[COL_OUT.bar]);
    const n = normName_(r[COL_OUT.name]);
    if (b && !barMap.has(b))  barMap.set(b, i);
    if (n && !nameMap.has(n)) nameMap.set(n, i);
  });

  const toAppend  = [];
  const updated   = [];   // names/barcodes auto-updated (no conflict)
  const conflicts = [];   // barcode-change / name-change rows for sidebar review

  rawData.forEach((rowIn, i) => {
    const rowNumber = FIRST_ROW + i;
    const bar  = rowIn[COL_IN.bar];
    const name = rowIn[COL_IN.name];
    if (!bar && !name) return;
    if (duplicateRowNumbers.has(rowNumber)) return; // held back until the duplicate is resolved

    const bKey = normBar_(bar);
    const nKey = normName_(name);

    const barIdx  = bKey ? (barMap.has(bKey)  ? barMap.get(bKey)  : null) : null;
    const nameIdx = nKey ? (nameMap.has(nKey) ? nameMap.get(nKey) : null) : null;

    // 1) Direct barcode match
    if (barIdx != null) {
      const out = prodData[barIdx];
      const existingNameKey = normName_(out[COL_OUT.name]);

      // 1a) Barcode matches, but the incoming name is different -> Name Change conflict
      if (nKey && existingNameKey && nKey !== existingNameKey) {
        conflicts.push({
          type:      'name',
          name:      String(out[COL_OUT.name]),   // stored (old) name
          newName:   String(name),                // incoming (new) name
          storedBar: String(out[COL_OUT.bar]),
          newBar:    String(bar),                  // unchanged, shown for context
          storedNpp: out[COL_OUT.npp],
          newNpp:    rowIn[COL_IN.npp],
          storedMrp: out[COL_OUT.mrp],
          newMrp:    rowIn[COL_IN.mrp]
        });
        return;
      }

      // 1b-i) Same digits but the leading zeros differ -> Barcode Change for review.
      // normBar_ deliberately ignores zeros so old and new forms still match,
      // but the difference is real and you should get to approve it.
      const storedBarRaw = String(out[COL_OUT.bar]).trim();
      const newBarRaw    = String(bar).trim();
      if (newBarRaw && storedBarRaw && newBarRaw !== storedBarRaw) {

        // Zero-only difference, and only when zeros are being ADDED:
        // apply it directly if the toggle is on. Losing zeros always asks.
        const gainingZeros = zeroOnlyDiff_(storedBarRaw, newBarRaw)
                             && newBarRaw.length > storedBarRaw.length;

        if (gainingZeros && zeroAutoIsOn_()) {
          out[COL_OUT.bar] = newBarRaw;
          if (changed_(rowIn[COL_IN.npp], out[COL_OUT.npp])) out[COL_OUT.npp] = rowIn[COL_IN.npp];
          if (changed_(rowIn[COL_IN.mrp], out[COL_OUT.mrp])) out[COL_OUT.mrp] = rowIn[COL_IN.mrp];
          updated.push(storedBarRaw + " \u2192 " + newBarRaw);
          return;
        }

        conflicts.push({
          type:      'barcode',
          name:      String(out[COL_OUT.name]),
          storedBar: storedBarRaw,
          newBar:    newBarRaw,
          storedNpp: out[COL_OUT.npp],
          newNpp:    rowIn[COL_IN.npp],
          storedMrp: out[COL_OUT.mrp],
          newMrp:    rowIn[COL_IN.mrp]
        });
        return;
      }

      // 1b-ii) Barcode identical -> compare NPP & MRP only, as before
      let hasChanged = false;
      if (changed_(rowIn[COL_IN.npp], out[COL_OUT.npp])) { out[COL_OUT.npp] = rowIn[COL_IN.npp]; hasChanged = true; }
      if (changed_(rowIn[COL_IN.mrp], out[COL_OUT.mrp])) { out[COL_OUT.mrp] = rowIn[COL_IN.mrp]; hasChanged = true; }
      if (hasChanged) updated.push(String(bar || name));
      return;
    }

    // 2) Barcode is new but name matches -> Barcode Change conflict
    if (nameIdx != null && bKey) {
      const ex = prodData[nameIdx];
      conflicts.push({
        type:      'barcode',
        name:      String(ex[COL_OUT.name]),
        storedBar: String(ex[COL_OUT.bar]),
        newBar:    String(bar),
        storedNpp: ex[COL_OUT.npp],
        newNpp:    rowIn[COL_IN.npp],
        storedMrp: ex[COL_OUT.mrp],
        newMrp:    rowIn[COL_IN.mrp]
      });
      return;
    }

    // 2b) Name matches but no new barcode (blank barcode) -> just update NPP/MRP in place
    if (nameIdx != null && !bKey) {
      const out = prodData[nameIdx];
      let hasChanged = false;
      if (changed_(rowIn[COL_IN.npp], out[COL_OUT.npp])) { out[COL_OUT.npp] = rowIn[COL_IN.npp]; hasChanged = true; }
      if (changed_(rowIn[COL_IN.mrp], out[COL_OUT.mrp])) { out[COL_OUT.mrp] = rowIn[COL_IN.mrp]; hasChanged = true; }
      if (hasChanged) updated.push(String(name));
      return;
    }

    // 3) Genuinely new -> append
    const newRow = Array(prodCols).fill("");
    newRow[COL_OUT.bar]    = bar;
    newRow[COL_OUT.name]   = name;
    newRow[COL_OUT.brand]  = rowIn[COL_IN.brand];
    newRow[COL_OUT.npp]    = rowIn[COL_IN.npp];
    newRow[COL_OUT.mrp]    = rowIn[COL_IN.mrp];
    newRow[COL_OUT.shelf]  = rowIn[COL_IN.shelf];
    newRow[COL_OUT.rating] = rowIn[COL_IN.rating];
    toAppend.push(newRow);
  });

  // --- WRITE UPDATES (protecting column H) ---
  if (prodData.length) {
    const blockBG = prodData.map(r => r.slice(1, 7));   // B..G
    productSheet.getRange(2, 2, blockBG.length, 6).setValues(blockBG);

    const blockI = prodData.map(r => [r[8]]);           // I
    productSheet.getRange(2, 9, blockI.length, 1).setValues(blockI);
  }

  // --- APPEND NEW ---
  if (toAppend.length) {
    const appendBG = toAppend.map(r => r.slice(1, 7));
    productSheet.getRange(lastProdRow + 1, 2, appendBG.length, 6).setValues(appendBG);

    const appendI = toAppend.map(r => [r[8]]);
    productSheet.getRange(lastProdRow + 1, 9, appendI.length, 1).setValues(appendI);
  }

  // --- SIDEBAR: barcode-change / name-change / duplicate conflicts, all together ---
  const allConflicts = conflicts.concat(duplicateConflicts);
  if (allConflicts.length) {
    showProductConflictSidebar_(allConflicts);
  }

  const barcodeChangeCount = conflicts.filter(c => c.type === 'barcode').length;
  const nameChangeCount    = conflicts.filter(c => c.type === 'name').length;
  const dupRowCount        = duplicateRowNumbers.size;

  ui.alert(
    "Done.\n" +
    "Auto-updated (matched): " + updated.length + "\n" +
    "Added (new): " + toAppend.length + "\n" +
    "Barcode changes to review: " + barcodeChangeCount + "\n" +
    "Name changes to review: " + nameChangeCount + "\n" +
    "Duplicate rows in Import List: " + dupRowCount +
    (dupRowCount ? " (fix these first, see sidebar)" : "") +
    (allConflicts.length ? "  →  see the sidebar on the right." : "")
  );
}

/**
 * Called from the sidebar's "Apply Selected" button.
 * `selected` is an array of { type, name, newName?, newBar, storedBar, newNpp, newMrp }.
 *   type === 'barcode' -> matched by (unchanged) name, write the new barcode.
 *   type === 'name'    -> matched by (unchanged) barcode, write the new name.
 * Re-finds each product in the (possibly changed) Product List and applies
 * the change, plus NPP/MRP if they differ.
 */
function applyProductConflicts(selected) {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const productSheet = ss.getSheetByName('Product List');
  if (!productSheet || !selected || !selected.length) return "Nothing to apply.";

  const lastProdRow = productSheet.getLastRow();
  const prodCols    = Math.max(productSheet.getLastColumn(), 9);
  const prodRows    = Math.max(0, lastProdRow - 1);
  if (prodRows <= 0) return "Product List is empty.";

  const prodData = productSheet.getRange(2, 1, prodRows, prodCols).getValues();
  const COL = { bar:1, name:2, npp:4, mrp:5 };

  // Two lookup maps: by name (for barcode-change conflicts) and by barcode (for name-change conflicts)
  const nameMap = new Map();
  const barMap  = new Map();
  prodData.forEach((r, i) => {
    const n = normName_(r[COL.name]);
    const b = normBar_(r[COL.bar]);
    if (n && !nameMap.has(n)) nameMap.set(n, i);
    if (b && !barMap.has(b))  barMap.set(b, i);
  });

  let applied = 0;
  selected.forEach(c => {
    // Barcode-change: prefer finding by the stored barcode (exact target row),
    // falling back to the name if that row has since moved or changed.
    const idx = c.type === 'name'
      ? barMap.get(normBar_(c.storedBar))
      : (barMap.has(normBar_(c.storedBar))
          ? barMap.get(normBar_(c.storedBar))
          : nameMap.get(normName_(c.name)));

    if (idx == null) return;

    if (c.type === 'name') {
      prodData[idx][COL.name] = c.newName;
    } else {
      prodData[idx][COL.bar] = c.newBar;
    }
    if (changed_(c.newNpp, prodData[idx][COL.npp])) prodData[idx][COL.npp] = c.newNpp;
    if (changed_(c.newMrp, prodData[idx][COL.mrp])) prodData[idx][COL.mrp] = c.newMrp;
    applied++;
  });

  // Write B..G only (barcode/name/npp/mrp all sit inside this block; H stays untouched)
  const blockBG = prodData.map(r => r.slice(1, 7));
  productSheet.getRange(2, 2, blockBG.length, 6).setValues(blockBG);

  return "Applied " + applied + " change(s).";
}

/**
 * Called from the sidebar's per-row "Delete" button on a Duplicate card.
 * Re-verifies the row's current content against the signature captured at
 * scan time before deleting — if the sheet changed underneath the sidebar
 * (manual edit, another delete, etc.), nothing is deleted and the caller is
 * told to re-scan, same safety convention as applyDuplicateResolution().
 */
function deleteImportRow(rowNumber, sig) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName('Import List');
  if (!importSheet) return { ok:false, msg:"Import List sheet not found." };

  const n = Number(rowNumber);
  if (!n || n < 4) return { ok:false, msg:"Invalid row number." };
  if (n > importSheet.getLastRow()) return { ok:false, msg:"That row no longer exists." };

  const rowVals     = importSheet.getRange(n, 1, 1, 9).getValues()[0];
  // Barcode must be read the same way it was at scan time (display value),
  // or the signature never matches and every delete is refused.
  rowVals[0]        = importSheet.getRange(n, 1, 1, 1).getDisplayValues()[0][0];
  const currentSig  = importRowSig_(rowVals);
  if (currentSig !== sig) {
    return { ok:false, msg:"That row's content changed since the scan. Nothing was deleted — please re-run Update Product List." };
  }

  importSheet.deleteRow(n);
  return { ok:true, msg:"Row " + n + " deleted." };
}

/* ---------------- helpers ---------------- */

function normBar_(v) {
  if (v === null || v === undefined) return "";
  const s = String(v).trim();
  // "003764" and 3764 both collapse to "3764" for MATCHING purposes only;
  // the padded string is still what gets stored.
  return /^\d+$/.test(s) ? s.replace(/^0+/, "") : s;
}

function normName_(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim().toUpperCase();
}

function normPrice_(v) {
  if (v === null || v === undefined || v === "") return "";
  const num = Number(v);
  return isNaN(num) ? String(v).trim() : String(num);
}

// import value counts as a change only if it's non-blank AND differs from stored
function changed_(importVal, storedVal) {
  if (importVal === "" || importVal === null || importVal === undefined) return false;
  const ni = Number(importVal), ns = Number(storedVal);
  if (!isNaN(ni) && !isNaN(ns) && String(storedVal).trim() !== "") return ni !== ns;
  return String(importVal).trim() !== String(storedVal).trim();
}

// Content signature for one Import List row (columns A..G) — used to verify
// a row hasn't changed between scan time and delete time.
function importRowSig_(rowArr) {
  return [rowArr[0], rowArr[1], rowArr[2], rowArr[3], rowArr[4], rowArr[5], rowArr[6]]
    .map(v => (v === null || v === undefined) ? "" : String(v).trim())
    .join("\u0001");
}

function showProductConflictSidebar_(conflicts) {
  const json = JSON.stringify(conflicts);
  const html =
'<!DOCTYPE html><html><head><base target="_top">' +
'<style>' +
'body{font-family:Arial,Helvetica,sans-serif;font-size:13px;margin:8px;}' +
'h3{margin:4px 0 8px;}' +
'.bar{margin:6px 0 10px;}' +
'.item{border:1px solid #ddd;border-radius:6px;padding:8px;margin-bottom:8px;position:relative;}' +
'.tag{display:inline-block;font-size:10px;font-weight:bold;text-transform:uppercase;letter-spacing:.03em;padding:2px 6px;border-radius:10px;margin-bottom:6px;}' +
'.tag.barcode{background:#e8f0fe;color:#1a73e8;}' +
'.tag.name{background:#fef7e0;color:#b06000;}' +
'.tag.duplicate{background:#fce8e6;color:#c5221f;}' +
'.nm{font-weight:bold;margin-bottom:4px;}' +
'.row{font-size:12px;color:#333;}' +
'.row.muted{color:#888;}' +
'.explain{font-size:11px;color:#555;margin:2px 0 6px;line-height:1.35;}' +
'.intro{font-size:12px;color:#555;margin:0 0 10px;line-height:1.4;}' +
'.old{color:#b00;}.new{color:#0a0;font-weight:bold;}' +
'button{cursor:pointer;padding:6px 10px;border-radius:6px;border:1px solid #888;background:#f5f5f5;}' +
'button.primary{background:#1a73e8;color:#fff;border-color:#1a73e8;}' +
'.muted{color:#777;font-size:11px;}' +
'.item.done{background:#f1f3f4;border-color:#cdd2d6;opacity:.7;}' +
'.item.done .nm{text-decoration:line-through;color:#777;}' +
'.tick{color:#0a0;font-weight:bold;margin-right:4px;}' +
'.donebox{background:#e6f4ea;border:1px solid #b7e1c1;color:#137333;padding:8px;border-radius:6px;font-size:12px;}' +
'button:disabled{cursor:default;opacity:.5;}' +
'.duprow{display:flex;justify-content:space-between;align-items:center;border-top:1px solid #eee;padding:5px 0;font-size:12px;gap:8px;}' +
'.duprow-info{flex:1;line-height:1.3;}' +
'.delbtn{background:#fce8e6;color:#c5221f;border:1px solid #f2b8b5;border-radius:6px;padding:3px 8px;font-size:11px;white-space:nowrap;}' +
'.resolved{color:#137333;font-weight:bold;font-size:12px;}' +
'.modal-overlay{position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.45);display:flex;align-items:center;justify-content:center;z-index:1000;}' +
'.modal-box{background:#fff;border-radius:8px;padding:16px;width:250px;box-shadow:0 4px 18px rgba(0,0,0,.25);}' +
'.modal-title{font-weight:bold;font-size:13px;margin-bottom:8px;}' +
'.modal-body{font-size:12px;color:#333;margin-bottom:14px;line-height:1.4;}' +
'.modal-actions{display:flex;justify-content:flex-end;gap:8px;}' +
'.modal-actions button.danger{background:#d93025;color:#fff;border-color:#d93025;}' +
'</style></head><body>' +
'<h3>Product Change Review</h3>' +
'<p class="intro">Some rows you imported already exist in the Product List, but something looks different. Check each one below, tick the ones you want to update, then click Apply selected. Duplicate rows have their own Delete button instead.</p>' +
'<div class="bar">' +
'<button id="selAll" onclick="setAll(true)">Select all</button> ' +
'<button id="clr" onclick="setAll(false)">Clear</button></div>' +
'<div id="list"></div>' +
'<div class="bar"><button id="applyBtn" class="primary" onclick="apply()">Apply selected</button></div>' +
'<div id="status" class="muted"></div>' +
'<script>' +
'var CONFLICTS=' + json + ';' +
'function esc(s){return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}' +
'function render(){' +
'var h="";' +
'CONFLICTS.forEach(function(c,i){' +
'if(c.type==="duplicate"){' +
'var isDone=c.rows.length<=1;' +
'var tag2=c.subtype==="barcode"?"<span class=\\"tag duplicate\\">Duplicate Barcode</span>":"<span class=\\"tag duplicate\\">Duplicate Name</span>";' +
'var explain2=c.subtype==="barcode"?' +
'("These rows in your Import List all use the same barcode (\\""+esc(c.key)+"\\"). Keep the one you want and delete the rest, then run Update Product List again.")' +
':' +
'("These rows in your Import List share the same product name and the same NPP/MRP (\\""+esc(c.key)+"\\") \\u2014 likely the same item entered twice. Keep the one you want and delete the rest, then run Update Product List again.")' +
';' +
'var body="";' +
'if(isDone){' +
'body="<div class=\\"resolved\\">&#10003; Resolved</div>";' +
'}else{' +
'c.rows.forEach(function(r,ri){' +
'var extras=(r.brand?(" ("+esc(r.brand)+")"):"")+' +
'((r.npp!==""&&r.npp!=null)?(" | NPP "+esc(r.npp)):"")+' +
'((r.mrp!==""&&r.mrp!=null)?(" | MRP "+esc(r.mrp)):"");' +
'body+="<div class=\\"duprow\\">"+' +
'"<div class=\\"duprow-info\\">Row "+r.rowNumber+": <b>"+esc(r.itemCode)+"</b> \\u2014 "+esc(r.productName)+extras+"</div>"+' +
'"<button class=\\"delbtn\\" onclick=\\"deleteRow("+i+","+ri+",this)\\">Delete</button>"+' +
'"</div>";' +
'});' +
'}' +
'h+="<div class=\\"item"+(isDone?" done":"")+"\\" id=\\"card"+i+"\\">"+tag2+' +
'"<div class=\\"explain\\">"+explain2+"</div>"+body+"</div>";' +
'return;' +
'}' +
'var isName=c.type==="name";' +
'var tag=isName?"<span class=\\"tag name\\">Name Change</span>":"<span class=\\"tag barcode\\">Barcode Change</span>";' +
'var headerName=isName?c.newName:c.name;' +
'var explain=isName?' +
'"This barcode is already in your Product List, but saved under a different name. Tick the box to rename it to what you just imported."' +
':' +
'"This product name is already in your Product List, but saved under a different barcode. Tick the box to update it to the new barcode."' +
';' +
'var mainRow=isName?' +
'("<div class=\\"row\\">Name: <span class=\\"old\\">"+esc(c.name)+"</span> &rarr; <span class=\\"new\\">"+esc(c.newName)+"</span></div>"+' +
'"<div class=\\"row muted\\">Barcode: "+esc(c.storedBar)+"</div>")' +
':' +
'("<div class=\\"row\\">Barcode: <span class=\\"old\\">"+esc(c.storedBar)+"</span> &rarr; <span class=\\"new\\">"+esc(c.newBar)+"</span></div>");' +
'var npp=(c.newNpp!==""&&c.newNpp!=null&&String(c.newNpp)!==String(c.storedNpp))?("<div class=\\"row\\">NPP: <span class=\\"old\\">"+esc(c.storedNpp)+"</span> &rarr; <span class=\\"new\\">"+esc(c.newNpp)+"</span></div>"):"";' +
'var mrp=(c.newMrp!==""&&c.newMrp!=null&&String(c.newMrp)!==String(c.storedMrp))?("<div class=\\"row\\">MRP: <span class=\\"old\\">"+esc(c.storedMrp)+"</span> &rarr; <span class=\\"new\\">"+esc(c.newMrp)+"</span></div>"):"";' +
'h+="<div class=\\"item\\" id=\\"card"+i+"\\">"+tag+' +
'"<div class=\\"explain\\">"+explain+"</div>"+' +
'"<label><br><input type=\\"checkbox\\" id=\\"cb"+i+"\\"> <span class=\\"nm\\">"+esc(headerName)+"</span></label>"+' +
'mainRow+npp+mrp+"</div>";' +
'});' +
'document.getElementById("list").innerHTML=h;' +
'}' +
'function setAll(v){CONFLICTS.forEach(function(c,i){var e=document.getElementById("cb"+i);if(e&&!e.disabled)e.checked=v;});}' +
'function markDone(idxs){idxs.forEach(function(i){' +
'var cb=document.getElementById("cb"+i);if(cb){cb.checked=false;cb.disabled=true;}' +
'var card=document.getElementById("card"+i);if(card){card.className="item done";' +
'var nm=card.querySelector(".nm");if(nm)nm.innerHTML="<span class=\\"tick\\">&#10003;</span>"+nm.innerHTML;}' +
'});}' +
'function remaining(){var n=0;CONFLICTS.forEach(function(c,i){var e=document.getElementById("cb"+i);if(e&&!e.disabled)n++;});return n;}' +
'function apply(){' +
'var sel=[],idxs=[];CONFLICTS.forEach(function(c,i){var e=document.getElementById("cb"+i);if(e&&!e.disabled&&e.checked){' +
'sel.push({type:c.type,name:c.name,newName:c.newName,storedBar:c.storedBar,newBar:c.newBar,newNpp:c.newNpp,newMrp:c.newMrp});idxs.push(i);}});' +
'if(!sel.length){document.getElementById("status").innerText="Nothing selected.";return;}' +
'var btn=document.getElementById("applyBtn");btn.disabled=true;' +
'document.getElementById("status").innerText="Applying...";' +
'google.script.run.withSuccessHandler(function(msg){' +
'markDone(idxs);' +
'var left=remaining();' +
'var st=document.getElementById("status");' +
'if(left===0){st.className="donebox";st.innerHTML="&#10003; "+msg+" All done \\u2014 you can close this panel.";btn.style.display="none";}' +
'else{st.className="muted";st.innerText=msg+" "+left+" still pending.";btn.disabled=false;}' +
'}).withFailureHandler(function(err){var st=document.getElementById("status");st.className="muted";st.innerText="Error: "+err.message;btn.disabled=false;}).applyProductConflicts(sel);' +
'}' +
'function showConfirm(title,body,onConfirm){' +
'var overlay=document.createElement("div");overlay.className="modal-overlay";' +
'overlay.innerHTML="<div class=\\"modal-box\\"><div class=\\"modal-title\\">"+title+"</div>"+' +
'"<div class=\\"modal-body\\">"+body+"</div>"+' +
'"<div class=\\"modal-actions\\"><button id=\\"modalCancel\\">Cancel</button> "+' +
'"<button id=\\"modalConfirm\\" class=\\"danger\\">Delete</button></div></div>";' +
'document.body.appendChild(overlay);' +
'document.getElementById("modalCancel").onclick=function(){document.body.removeChild(overlay);};' +
'document.getElementById("modalConfirm").onclick=function(){document.body.removeChild(overlay);onConfirm();};' +
'}' +
'function deleteRow(ci,ri,btn){' +
'var r=CONFLICTS[ci].rows[ri];' +
'showConfirm(' +
'"Delete this row?",' +
'"Row "+r.rowNumber+": <b>"+esc(r.itemCode)+"</b> \\u2014 "+esc(r.productName)+"<br>This removes it from the Import List. This cannot be undone.",' +
'function(){ doDeleteRow(ci,ri,btn); }' +
');' +
'}' +
'function doDeleteRow(ci,ri,btn){' +
'var r=CONFLICTS[ci].rows[ri];' +
'var rowNumber=r.rowNumber, sig=r.sig;' +
'btn.disabled=true;btn.innerText="Deleting...";' +
'google.script.run.withSuccessHandler(function(res){' +
'if(res&&res.ok){' +
'CONFLICTS[ci].rows.splice(ri,1);' +
'CONFLICTS.forEach(function(c){if(c.type==="duplicate"){c.rows.forEach(function(rr){if(rr.rowNumber>rowNumber)rr.rowNumber=rr.rowNumber-1;});}});' +
'render();' +
'document.getElementById("status").className="muted";' +
'document.getElementById("status").innerText=res.msg;' +
'}else{' +
'btn.disabled=false;btn.innerText="Delete";' +
'document.getElementById("status").className="muted";' +
'document.getElementById("status").innerText=(res&&res.msg)?res.msg:"Failed.";' +
'}' +
'}).withFailureHandler(function(err){' +
'btn.disabled=false;btn.innerText="Delete";' +
'document.getElementById("status").innerText="Error: "+err.message;' +
'}).deleteImportRow(rowNumber,sig);' +
'}' +
'render();' +
'</script></body></html>';

  const out = HtmlService.createHtmlOutput(html).setTitle('Product Change Review').setWidth(320);
  SpreadsheetApp.getUi().showSidebar(out);
}
/* ---------------- Auto-apply zero-only barcode changes ---------------- */

const ZERO_AUTO_KEY_ = 'AUTO_APPLY_ZERO_BARCODE';

// Default is ON — zero-only differences apply straight away.
function zeroAutoIsOn_() {
  const v = PropertiesService.getDocumentProperties().getProperty(ZERO_AUTO_KEY_);
  return v === null ? true : v === 'true';
}

function toggleZeroAuto() {
  const now = !zeroAutoIsOn_();
  PropertiesService.getDocumentProperties()
    .setProperty(ZERO_AUTO_KEY_, String(now));

  SpreadsheetApp.getUi().alert(
    "Auto-apply zero-only barcode changes: " + (now ? "ON" : "OFF") + "\n\n" +
    (now
      ? "When the only difference is leading zeros (316709 \u2192 00316709), the "
        + "Product List is updated straight away without asking."
      : "Zero-only differences will be shown in the sidebar for you to approve, "
        + "like any other barcode change.") +
    "\n\nThe menu label updates next time the file is opened."
  );
}

// True when two barcodes are the same digits and differ only by leading zeros.
function zeroOnlyDiff_(a, b) {
  const s1 = String(a).trim(), s2 = String(b).trim();
  if (s1 === s2) return false;
  if (!/^\d+$/.test(s1) || !/^\d+$/.test(s2)) return false;
  return s1.replace(/^0+/, "") === s2.replace(/^0+/, "");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Product Tools')
    .addItem('Add many stickers…', 'openAddMany')
    .addItem('Place ADD MANY button on Print Console', 'placeAddManyButton')
    .addSeparator()
    .addItem('Update Product List (run import)', 'updateProductList2')
    .addSeparator()
    .addItem('Find Duplicates in Product List', 'executeFindDuplicates')
    .addSeparator()
    .addItem('Export Products to XLS', 'exportProductsToXlsx')
    .addSeparator()
    .addItem(
      'Auto-apply zero-only barcode changes: ' + (zeroAutoIsOn_() ? 'ON' : 'OFF'),
      'toggleZeroAuto'
    )
    .addToUi();
}