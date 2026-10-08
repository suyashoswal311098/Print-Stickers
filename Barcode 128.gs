/***** OWN BARCODE (Code 128) — no website needed *****
 * code128Text_(text) turns a code like "49876602" or "88008990-0.83" into the characters that the
 * free Google font "Libre Barcode 128" draws as a real, scannable Code 128 barcode
 * (start + data + checksum + stop). Digit runs use set C (2 digits per bar group = shorter barcode).
 * Use in a cell:  =BARCODE128(C7)  with the cell font set to "Libre Barcode 128".
 * code128WidthEm_(encoded) = barcode width in font-size units (each symbol 0.33 em, stop 0.39 em).
 */
function code128Values_(text) {
  const s = String(text == null ? '' : text).replace(/[^\x20-\x7E]/g, '');
  const vals = [];
  let set = null, i = 0;
  const isD = ch => ch >= '0' && ch <= '9';
  const runAt = j => { let k = j; while (k < s.length && isD(s[k])) k++; return k - j; };
  const to = x => {
    if (set === x) return;
    vals.push(set === null ? (x === 'C' ? 105 : 104) : (x === 'C' ? 99 : 100));
    set = x;
  };
  const b = ch => { to('B'); vals.push(ch.charCodeAt(0) - 32); };
  while (i < s.length) {
    let run = runAt(i);
    if (run >= 4 || (i === 0 && run === s.length && run >= 2 && run % 2 === 0)) {
      if (run % 2) { b(s[i]); i++; run--; }
      to('C');
      for (let n = 0; n < run; n += 2) { vals.push(Number(s.substr(i, 2))); i += 2; }
    } else {
      b(s[i]); i++;
    }
  }
  if (!vals.length) return [];
  let sum = vals[0];
  for (let k = 1; k < vals.length; k++) sum += vals[k] * k;
  vals.push(sum % 103, 106);
  return vals;
}

function code128Text_(text) {
  return code128Values_(text)
    .map(v => String.fromCharCode(v === 0 ? 194 : v < 95 ? v + 32 : v + 100))
    .join('');
}

function code128WidthEm_(encoded) {
  const n = String(encoded || '').length;
  return n ? (n - 1) * 0.33 + 0.39 : 0;
}

/**
 * Own barcode for a cell: =BARCODE128(C7) and set the cell's font to "Libre Barcode 128".
 * @param {string} text the code
 * @return the barcode characters
 * @customfunction
 */
function BARCODE128(text) {
  return code128Text_(text);
}

// ---- Used by LOAD (load2 in Load Stickers 3) ------------------------------------------------
// true = draw the barcode with our own font; false = keep the template's =image(...) website barcode.
const OWN_BARCODE_ = true;

// Finds the barcode cell in the sticker template (the cell whose formula fetches a barcode image)
// and the size of that space on the sticker sheet. Returns null if not found (then nothing changes).
function ownBarcodeSetup_(rstk, stkr) {
  if (!OWN_BARCODE_) return null;
  // Only the barcode picture formula (…barcode/image?content=…&symbology=…), never the logo
  const f = rstk.getRange('A1:F9').getFormulas();
  let r = 0, c = 0, n = 0;
  for (let i = 0; i < f.length; i++)
    for (let j = 0; j < f[i].length; j++)
      if (/symbology=|barcode\/image/i.test(f[i][j])) { if (!r) { r = i + 1; c = j + 1; } n++; }
  if (n !== 1) {   // none, or more than one → don't guess
    console.warn('Own barcode: found ' + n + ' barcode formula(s) in the template, using the website barcode');
    return null;
  }
  const m = rstk.getRange(r, c).getMergedRanges();
  const rows = m.length ? m[0].getNumRows() : 1, cols = m.length ? m[0].getNumColumns() : 1;
  const w = [0, 6, 12].map(off => {
    let s = 0;
    for (let k = 0; k < cols; k++) s += stkr.getColumnWidth(off + c + k);
    return s;
  });
  let h = 0;
  const hs = [];
  for (let k = 0; k < rows; k++) { hs.push(stkr.getRowHeight(r + k)); h += hs[k]; }
  return { r: r, c: c, w: w, h: h, hs: hs };
}

// Writes the barcode for one sticker. top = sticker's first row, colOffset = 0 / 6 / 12, stc = 1..3.
function ownBarcodePut_(sheet, bc, top, colOffset, stc, code) {
  const enc = code128Text_(code);
  if (!enc) return;
  const em = code128WidthEm_(enc);
  // The font's bars are 0.59 em tall from the baseline; below them is 0.4 em of empty "descender" space.
  // Fill the width (94%) and let the bars reach the full height: if the text box is taller than the cell,
  // align it to the TOP so only the empty space below the bars is cut off.
  const px = Math.min(bc.w[stc - 1] * 0.94 / em, bc.h / 0.6);
  const pt = Math.max(6, Math.floor(px * 0.75));                  // px → pt
  sheet.getRange(top + bc.r - 1, colOffset + bc.c)
    .setValue(enc)
    .setFontFamily('Libre Barcode 128')
    .setFontSize(pt)
    .setFontColor('#000000')
    .setHorizontalAlignment('center')
    .setVerticalAlignment(pt / 0.75 > bc.h ? 'top' : 'middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  // Keep the sticker's row heights exactly as they are (a big font must not stretch the rows)
  if (stc === 1) bc.hs.forEach((hh, k) => sheet.setRowHeightsForced(top + bc.r - 1 + k, 1, hh));
}
