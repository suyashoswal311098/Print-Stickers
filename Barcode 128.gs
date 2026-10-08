/***** OWN BARCODE (Code 128) — no website needed *****
 * code128Text_(text) turns a code like "49876602" or "88008990-0.83" into the characters that the
 * free Google font "Libre Barcode 128" draws as a real, scannable Code 128 barcode
 * (start + data + checksum + stop). Digit runs use set C (2 digits per bar group = shorter barcode).
 * Use in a cell:  =BARCODE128(C7)  with the cell font set to "Libre Barcode 128".
 * LOAD uses a barcode PICTURE instead (code128Png_), the font is only the fallback.
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
// Main way: a barcode PICTURE made here (code128Png_) put INSIDE the cell, drawn in the cell's exact
// shape so it fills the whole barcode space. If Google refuses the picture: the font way (smaller).
function ownBarcodePut_(sheet, bc, top, colOffset, stc, code) {
  if (!String(code || '').trim()) return;
  const cell = sheet.getRange(top + bc.r - 1, colOffset + bc.c);
  const w = bc.w[stc - 1], h = bc.h;
  bc.cache = bc.cache || {};
  const key = code + '|' + w + 'x' + h;
  try {
    if (!bc.cache[key]) {
      const mods = code128Modules_(code).length + 2 * 6;              // 6 white modules each side
      const scale = Math.max(2, Math.ceil(600 / mods));               // sharp: ~600+ px wide
      const height = Math.max(20, Math.round(mods * scale * h / w));   // same shape as the cell
      const png = code128Png_(code, scale, height, 6);
      const b64 = Utilities.base64Encode(png.map(x => x > 127 ? x - 256 : x));
      bc.cache[key] = SpreadsheetApp.newCellImage()
        .setSourceUrl('data:image/png;base64,' + b64)
        .setAltTextTitle(String(code))
        .build();
    }
    cell.setValue(bc.cache[key]);
  } catch (err) {
    console.warn('Own barcode picture failed (' + err.message + '), using the font');
    const enc = code128Text_(code);
    const px = Math.min(w * 0.92 / code128WidthEm_(enc), h * 0.98);
    cell.setValue(enc)
      .setFontFamily('Libre Barcode 128')
      .setFontSize(Math.max(6, Math.floor(px * 0.75)))
      .setFontColor('#000000')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
  // Keep the sticker's row heights exactly as they are
  if (stc === 1) bc.hs.forEach((hh, k) => sheet.setRowHeightsForced(top + bc.r - 1 + k, 1, hh));
}

// ---- Barcode picture (PNG) made in the script: black/white, 1 bit per pixel, stored (uncompressed) zlib ----
// code128Pattern_(text) → array of bar widths in modules (bar, space, bar, …), from the Code 128 table.
const C128_PATTERNS_ = ('212222 222122 222221 121223 121322 131222 122213 122312 132212 221213 221312 231212 112232 122132 122231 113222 123122 123221 223211 221132 221231 213212 223112 312131 311222 321122 321221 312212 322112 322211 212123 212321 232121 111323 131123 131321 112313 132113 132311 211313 231113 231311 112133 112331 132131 113123 113321 133121 313121 211331 231131 213113 213311 213131 311123 311321 331121 312113 312311 332111 314111 221411 431111 111224 111422 121124 121421 141122 141221 112214 112412 122114 122411 142112 142211 241211 221114 413111 241112 134111 111242 121142 121241 114212 124112 124211 411212 421112 421211 212141 214121 412121 111143 111341 131141 114113 114311 411113 411311 113141 114131 311141 411131 211412 211214 211232 2331112').split(' ');

function code128Modules_(text) {
  const vals = code128Values_(text);
  const bits = [];
  vals.forEach(v => {
    const p = C128_PATTERNS_[v];
    for (let i = 0; i < p.length; i++) for (let k = 0; k < Number(p[i]); k++) bits.push(i % 2 === 0 ? 1 : 0);
  });
  return bits; // 1 = black module
}

function crc32_(bytes, start, end) {
  let c, crc = 0xFFFFFFFF;
  for (let n = start; n < end; n++) {
    c = (crc ^ bytes[n]) & 0xFF;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

// PNG bytes (0..255) of the barcode: every module `scale` px wide, `quiet` white modules each side, height px.
function code128Png_(text, scale, height, quiet) {
  const mods = code128Modules_(text);
  if (!mods.length) return null;
  const line = [];
  for (let q = 0; q < quiet; q++) line.push(0);
  mods.forEach(m => line.push(m));
  for (let q = 0; q < quiet; q++) line.push(0);
  const W = line.length * scale, H = height, rowLen = Math.ceil(W / 8) + 1;
  const row = new Array(rowLen).fill(0xFF); row[0] = 0;            // filter 0; 1 = white in grayscale 1-bit
  for (let x = 0; x < W; x++) if (line[Math.floor(x / scale)]) row[1 + (x >> 3)] &= ~(0x80 >> (x & 7)) & 0xFF;
  const raw = [];
  for (let y = 0; y < H; y++) for (let i = 0; i < rowLen; i++) raw.push(row[i]);
  // zlib: stored blocks (no compression) + Adler-32
  const z = [0x78, 0x01];
  for (let p = 0; p < raw.length || p === 0; p += 65535) {
    const len = Math.min(65535, raw.length - p), last = p + len >= raw.length ? 1 : 0;
    z.push(last, len & 0xFF, len >> 8, ~len & 0xFF, (~len >> 8) & 0xFF);
    for (let i = 0; i < len; i++) z.push(raw[p + i]);
    if (last) break;
  }
  let a = 1, b = 0;
  for (let i = 0; i < raw.length; i++) { a = (a + raw[i]) % 65521; b = (b + a) % 65521; }
  z.push(b >> 8, b & 0xFF, a >> 8, a & 0xFF);
  const out = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
  const u32 = (arr, v) => arr.push((v >>> 24) & 0xFF, (v >>> 16) & 0xFF, (v >>> 8) & 0xFF, v & 0xFF);
  const chunk = (type, data) => {
    u32(out, data.length);
    const s = out.length;
    for (let i = 0; i < 4; i++) out.push(type.charCodeAt(i));
    for (let i = 0; i < data.length; i++) out.push(data[i]);
    u32(out, crc32_(out, s, out.length));
  };
  const ihdr = []; u32(ihdr, W); u32(ihdr, H); ihdr.push(1, 0, 0, 0, 0);  // 1-bit grayscale
  chunk('IHDR', ihdr);
  chunk('IDAT', z);
  chunk('IEND', []);
  return out;
}
