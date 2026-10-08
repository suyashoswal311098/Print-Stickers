  /**
   * DUPLICATE CLEANUP
   * Finds products in 'Product List' that duplicate each other by:
   *   - BARCODE match (always — a re-barcoded item keeps the same identity), or
   *   - NAME match AND matching NPP+MRP (price) — same name alone is NOT enough,
   *     since many genuinely different products share a generic name. Only when
   *     the price also matches do we treat it as the same product re-entered.
   * Matches chain transitively (union-find) so linked rows form one group.
   * You pick which row to KEEP per group in a sidebar; the rest are deleted.
   *
   * Safe by design:
   *  - Default per group is "Keep all" (no deletion) — you must actively pick a keeper.
   *  - Deletion is confirmed with an exact count and is bottom-up (no index shifting).
   *  - Each row to delete is re-verified by content signature at apply time; if the
   *    sheet changed since the scan, nothing is deleted and you're told to re-scan.
   *
   * Assign executeFindDuplicates to a button (same pattern as your other buttons).
   */

  function executeFindDuplicates() {
    findDuplicates();
  }

  function findDuplicates() {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const ui    = SpreadsheetApp.getUi();
    const sheet = ss.getSheetByName('Product List');
    if (!sheet) { ui.alert("Product List not found."); return; }

    const lastRow = sheet.getLastRow();
    const cols    = Math.max(sheet.getLastColumn(), 9);
    const n       = Math.max(0, lastRow - 1);
    if (n <= 0) { ui.alert("No products to scan."); return; }

    const data = sheet.getRange(2, 1, n, cols).getValues();
    const C = { bar:1, name:2, brand:3, npp:4, mrp:5, shelf:6, rating:8 };

    // ---- union-find ----
    const parent = [];
    for (let i = 0; i < n; i++) parent[i] = i;
    function find(x){ while (parent[x] !== x){ parent[x] = parent[parent[x]]; x = parent[x]; } return x; }
    function union(a,b){ const ra = find(a), rb = find(b); if (ra !== rb) parent[ra] = rb; }

    const firstByNamePrice = new Map();
    const firstByBar       = new Map();
    for (let i = 0; i < n; i++) {
      const nm = dcNormName_(data[i][C.name]);
      const br = dcNormBar_(data[i][C.bar]);
      if (!nm && !br) continue;                 // fully blank row — ignore
      if (br) { if (firstByBar.has(br)) union(i, firstByBar.get(br)); else firstByBar.set(br, i); }

      const npp = dcNormPrice_(data[i][C.npp]);
      const mrp = dcNormPrice_(data[i][C.mrp]);
      if (nm && npp !== "" && mrp !== "") {       // only group by name when price is known AND matches
        const key = nm + "\u0001" + npp + "\u0001" + mrp;
        if (firstByNamePrice.has(key)) union(i, firstByNamePrice.get(key));
        else firstByNamePrice.set(key, i);
      }
    }

    // ---- collect components of size >= 2 ----
    const comp = new Map();   // root -> [rowIdx,...]
    for (let i = 0; i < n; i++) {
      const nm = dcNormName_(data[i][C.name]);
      const br = dcNormBar_(data[i][C.bar]);
      if (!nm && !br) continue;
      const r = find(i);
      if (!comp.has(r)) comp.set(r, []);
      comp.get(r).push(i);
    }

    const groups = [];
    comp.forEach(idxs => {
      if (idxs.length >= 2) {
        groups.push(idxs.map(i => ({
          row:  i + 2,                                  // physical sheet row
          bar:  String(data[i][C.bar]),
          name: String(data[i][C.name]),
          npp:  data[i][C.npp],
          mrp:  data[i][C.mrp],
          sig:  dcSig_(data[i], C)
        })));
      }
    });

    if (!groups.length) { ui.alert("No duplicates found."); return; }

    showDuplicateSidebar_(groups);
    ui.alert("Found " + groups.length + " duplicate group(s). Review them in the sidebar on the right.");
  }

  /**
   * Called from the sidebar. `deletes` = [{ row, sig }, ...].
   * Verifies each row still matches its captured signature, then deletes bottom-up.
   */
  function applyDuplicateResolution(deletes) {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Product List');
    if (!sheet || !deletes || !deletes.length) return { ok:false, msg:"Nothing to delete." };

    const lastRow = sheet.getLastRow();
    const cols    = Math.max(sheet.getLastColumn(), 9);
    const n       = Math.max(0, lastRow - 1);
    if (n <= 0) return { ok:false, msg:"Product List is empty." };

    const data = sheet.getRange(2, 1, n, cols).getValues();
    const C = { bar:1, name:2, brand:3, npp:4, mrp:5, shelf:6, rating:8 };

    // verify nothing shifted/changed since the scan
    let stale = false;
    deletes.forEach(d => {
      const ri = d.row - 2;
      if (ri < 0 || ri >= n || dcSig_(data[ri], C) !== d.sig) stale = true;
    });
    if (stale) return { ok:false, msg:"The sheet changed since the scan. Nothing was deleted — please re-run Find Duplicates." };

    // delete bottom-up (highest row first), de-duped
    const rows = deletes.map(d => d.row).sort((a,b) => b - a);
    const seen = {};
    let count = 0;
    rows.forEach(r => { if (!seen[r]) { seen[r] = 1; sheet.deleteRow(r); count++; } });

    return { ok:true, deleted: count };
  }

  /* ---------------- helpers (dc-prefixed to avoid clashes) ---------------- */

  function dcNormName_(v) {
    if (v === null || v === undefined) return "";
    return String(v).trim().toUpperCase();
  }

  function dcNormBar_(v) {
    if (v === null || v === undefined) return "";
    return String(v).trim();
  }

  function dcNormPrice_(v) {
    if (v === null || v === undefined || v === "") return "";
    const num = Number(v);
    return isNaN(num) ? String(v).trim() : String(num);
  }

  function dcSig_(rowArr, C) {
    return [rowArr[C.bar], rowArr[C.name], rowArr[C.brand], rowArr[C.npp], rowArr[C.mrp], rowArr[C.shelf], rowArr[C.rating]]
      .map(v => (v === null || v === undefined) ? "" : String(v).trim())
      .join("\u0001");
  }

  function showDuplicateSidebar_(groups) {
    const json = JSON.stringify(groups);
    const html =
  '<!DOCTYPE html><html><head><base target="_top">' +
  '<style>' +
  'body{font-family:Arial,Helvetica,sans-serif;font-size:13px;margin:8px;}' +
  'h3{margin:4px 0 8px;}' +
  '.group{border:1px solid #ddd;border-radius:6px;padding:8px;margin-bottom:10px;}' +
  '.ghead{font-weight:bold;margin-bottom:6px;}' +
  '.opt{display:block;font-size:12px;margin:4px 0;cursor:pointer;line-height:1.35;}' +
  '.skip{color:#555;}' +
  '.bar{margin:6px 0 10px;}' +
  'button{cursor:pointer;padding:6px 10px;border-radius:6px;border:1px solid #888;background:#f5f5f5;}' +
  'button.primary{background:#d93025;color:#fff;border-color:#d93025;}' +
  'button:disabled{cursor:default;opacity:.5;}' +
  '.muted{color:#777;font-size:11px;}' +
  '.donebox{background:#e6f4ea;border:1px solid #b7e1c1;color:#137333;padding:8px;border-radius:6px;font-size:12px;}' +
  '</style></head><body>' +
  '<h3>Duplicate Cleanup</h3>' +
  '<div class="muted bar">Pick one row to keep in each group. Anything left on &ldquo;Keep all&rdquo; is untouched.</div>' +
  '<div id="list"></div>' +
  '<div class="bar"><button id="applyBtn" class="primary" onclick="apply()">Delete selected duplicates</button></div>' +
  '<div id="status" class="muted"></div>' +
  '<script>' +
  'var GROUPS=' + json + ';' +
  'function esc(s){return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}' +
  'function detail(r){var p=[];p.push("barcode "+esc(r.bar||"\\u2014"));p.push(esc(r.name||"\\u2014"));if(r.npp!==""&&r.npp!=null)p.push("NPP "+esc(r.npp));if(r.mrp!==""&&r.mrp!=null)p.push("MRP "+esc(r.mrp));return p.join(" | ");}' +
  'function render(){var h="";GROUPS.forEach(function(g,gi){' +
  'h+="<div class=\\"group\\">";' +
  'h+="<div class=\\"ghead\\">Group "+(gi+1)+" \\u2014 "+g.length+" rows</div>";' +
  'h+="<label class=\\"opt skip\\"><input type=\\"radio\\" name=\\"g"+gi+"\\" value=\\"skip\\" checked> Keep all (don\\u2019t delete any)</label>";' +
  'g.forEach(function(r,ri){' +
  'h+="<label class=\\"opt\\"><input type=\\"radio\\" name=\\"g"+gi+"\\" value=\\""+ri+"\\"> <b>Keep</b> Row "+r.row+": "+detail(r)+"</label>";' +
  '});' +
  'h+="</div>";});' +
  'document.getElementById("list").innerHTML=h;}' +
  'function gather(){var dels=[];GROUPS.forEach(function(g,gi){' +
  'var sel=document.querySelector("input[name=\\"g"+gi+"\\"]:checked");if(!sel||sel.value==="skip")return;' +
  'var k=parseInt(sel.value,10);g.forEach(function(r,ri){if(ri!==k)dels.push({row:r.row,sig:r.sig});});});return dels;}' +
  'function apply(){var dels=gather();var st=document.getElementById("status");' +
  'if(!dels.length){st.className="muted";st.innerText="Nothing selected \\u2014 every group is set to Keep all.";return;}' +
  'if(!window.confirm("Delete "+dels.length+" row(s)? This cannot be undone."))return;' +
  'var btn=document.getElementById("applyBtn");btn.disabled=true;st.className="muted";st.innerText="Deleting...";' +
  'google.script.run.withSuccessHandler(function(res){' +
  'if(res&&res.ok){st.className="donebox";st.innerHTML="&#10003; Deleted "+res.deleted+" row(s). Close this panel and re-run Find Duplicates to verify.";btn.style.display="none";' +
  'var ins=document.querySelectorAll("input");for(var i=0;i<ins.length;i++)ins[i].disabled=true;}' +
  'else{st.className="muted";st.innerText=(res&&res.msg)?res.msg:"Failed.";btn.disabled=false;}' +
  '}).withFailureHandler(function(err){st.className="muted";st.innerText="Error: "+err.message;btn.disabled=false;}).applyDuplicateResolution(dels);}' +
  'render();' +
  '</script></body></html>';

    const out = HtmlService.createHtmlOutput(html).setTitle('Duplicate Cleanup').setWidth(340);
    SpreadsheetApp.getUi().showSidebar(out);
  }