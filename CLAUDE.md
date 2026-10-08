# Print Stickers — Apps Script mirror

This repo is a copy of the **Sticker Printing** Apps Script project. The live code is in the Apps Script editor; this repo is for history, reading and editing.

## Links
- Apps Script project: `1ooalz3QzT--oqUa32ra84qwkc9N2s1zgc3PZSZIR_Hos-Cd-iKgHGBLi`
  (https://script.google.com/home/projects/1ooalz3QzT--oqUa32ra84qwkc9N2s1zgc3PZSZIR_Hos-Cd-iKgHGBLi/edit)
- Bound sheet: "Print Stickers" `1z--0NDAIVD7K529-FzLj2LX_kxwUsVKi_lBeFNfy95U`
  (https://docs.google.com/spreadsheets/d/1z--0NDAIVD7K529-FzLj2LX_kxwUsVKi_lBeFNfy95U/edit)
- Deployments (8 Oct 2026): only the HEAD test deployment `AKfycbzz49zIoGRaZADqNfm7576adJtoCDrI8LaqcbhYg3A`; no versioned web app.

## Files
- Exact copy of the editor: script files as `.gs` (clasp names them `.js`), plus `appsscript.json`. Nothing else except `CLAUDE.md` and `.gitattributes`.
- `SheetJS.gs` = bundled SheetJS 0.18.5 library (~880 KB, used by `exportProductsToXlsx`). Do not edit.
- Manifest uses the advanced **Sheets v4** service and lists sheet macros (Reset Pages, Print, Clear Console, insert, Clear Import Product List, filters, many "Untitled macro" entries).
- Copied 8 Oct 2026 from a fresh `clasp clone` (logged in as erp.nutparadise@gmail.com), 23 files.

## Pushing changes (clasp)
1. Always start from a fresh `clasp clone` in a scratch folder (`clasp pull` does not delete removed/renamed files).
2. Before pushing, clone again and diff: if someone edited the editor in the meantime, merge first.
3. Edit, `clasp push`, re-clone and compare to confirm it landed.
4. Copy changed files back here (`.js` → `.gs`), commit and push to GitHub.

## Rules
- The user is not a coder: explain in plain language.
- Ask before any action that saves, renames or deletes real data in the sheet.
- clasp login in a cloud session: `clasp login --no-localhost` reading stdin from a fifo; the login dies with the session.

## ADD MANY (8 Oct 2026) — pushed to HEAD, not yet seen working
- User: pick many products at once instead of B2 one by one. `Add Many.gs` + `Add Many Dialog.html`: window with a search box (all typed words must be in the name), every match listed with a Qty box; typing a qty ticks it; ticks kept across searches; shows stickers already in the list (C5) + new vs the 240 limit; Submit → `amInsertMany` puts each product in B2 + qty in C2 + D2 = 1, flush, runs `insert3(null, true, cachedProdList)` (same as INSERT), lists skipped ones.
- Product names = B2's dropdown list (`amGetProducts`, same as `pullFromInventory`).
- Opened from a purple ADD MANY image button on Print Console (column H, row 3, left of LOAD; `placeAddManyButton` puts it there once, alt title "ADD MANY", script `openAddMany`; base64 PNG in the file) and Product Tools → "Add many stickers…".

## COPY TO IMPORT LIST (8 Oct 2026) — pushed to HEAD, not yet seen working
- User: put the Print Console sticker products into the Import List (for the XLS export they run themselves; no automatic export). `Stickers To Import List.gs` → `stickersToImportList`: Print Console rows 8+ (A code as displayed, B name, G brand, H NPP, I MRP) + Product List shelf life (G) / premium rating (I) → Import List A–G, ADDED below the last row with A/B filled (row 4 down; adds sheet rows if needed); one row per product; skips products already there (same barcode via `normBar_` or same name via `normName_`). Barcode column set to text; Premium Rating dropdown copied from Product List I2.
- Teal image button at Print Console A3 (`placeImportListButton`, alt title "COPY TO IMPORT LIST"). Menu: Product Tools → "Copy stickers to Import List", "Place buttons on Print Console" (`placePrintConsoleButtons` = both buttons, skips ones already there).
- **Fast ADD MANY (8 Oct 2026, pushed, not yet seen):** `amInsertMany` now looks up barcode (Product List B) + colour (H) by name (C, first match) and writes all rows at once (same row shape as `insert3`: A barcode, C qty, D 1, J colour value + background, others blank for the sheet's own formulas). Safety check: the first product is put in B2 and the sheet's A2/H2 must equal the lookup; else (or for names not found) the old slow per-item `insert3` path runs. Dialog shows "⚡" (fast) or "(slow way)".
- **Order before Submit (8 Oct 2026, pushed, checked with a Playwright stub, not yet seen live):** with the search box empty, the Ticked list shows every ticked product numbered in the order they go in, with a ⠿ drag handle (pointer listeners on `document`, row moved in the DOM, `order` rebuilt on release) and ↑ ↓ buttons. While searching, ticked rows have no order controls ("clear the search to change the order").
- **↕ Order button + barcode-prefixed names (8 Oct 2026, pushed):** button next to the search box clears the search and shows the ordered Ticked list. The B2 dropdown values look like "49876602 PUMPKIN SEEDS 100GM" (barcode + name), so the fast lookup also tries "<barcode> <name>" → Product List by barcode (name must match too); unknown → slow path.
- **Tick = qty 1 (8 Oct 2026, pushed):** clicking a row / its tick box ticks it with qty 1 and selects the qty box (typing replaces it); clicking a ticked row unticks it.
- **Smooth drag (8 Oct 2026, pushed, Playwright-checked):** the dragged row follows the pointer (transform), the others slide aside (0.18 s), it glides into place, then the list redraws. Mouse: press anywhere on a row (not qty/buttons) and move ≥5 px (a plain click still ticks/unticks; `noClick` stops the click after a drag). Touch: ⠿ handle only. Auto-scroll near the list edges.
- **Save order in Product List (8 Oct 2026, pushed, not yet seen live; user chose "both"):** tick box "Save this order in Product List" (default on, remembered per browser `am_saveOrd`). After a Submit with 2+ products, `amSaveOrder(names)` MOVES those products' whole rows in Product List (`moveRows`, upward only) into the Submit order as one block starting at the row of the first of them. `amGetProducts` sorts the window's list by Product List row, so searches show the saved order. Shared lookup `amProductIndex_` (name or "<barcode> <name>").
- **Cursor in search at once (8 Oct 2026, pushed):** the search box is enabled + focused as soon as the window opens (`focusQ`, retried at load / 150 ms / 500 ms); typing before the products arrive shows "Loading products…", then the matches.
- **Sharper buttons (8 Oct 2026, pushed):** user said the picture buttons look blurry. PNGs redrawn at 4x (328x152, shown 82x38). `amPlaceButton_` (Add Many.gs) places a button or, if one with the same alt title exists, swaps the picture in the same spot/size (keeps where the user dragged it). Run Product Tools → Place buttons on Print Console to refresh. Pictures are still bitmaps; a Google Drawing (Insert → Drawing, assign script `openAddMany` / `stickersToImportList`) is the only fully sharp option.
- **PKD + Pack (8 Oct 2026, pushed, not yet seen live):** PKD/EXP were blank because `insert3` only writes them when F2 (shelf life) is typed. Window now has a PKD date (default today, sent as `opts.pkd` 'yyyy-mm-dd') and a Pack box per product (default 1, `packs[name]`). Fast path writes E = PKD, F = EXP = PKD + Product List shelf life (G) when > 0, D = pack, E:F formatted DD/MM/YYYY. Slow path puts pack/PKD/shelf into D2/E2/F2 before `insert3`, then restores E2 (formula or value).

## Own barcode + sticker fixes (8 Oct 2026) — pushed to HEAD, not yet seen working
- LOAD = `load2` (`Load Stickers 3.gs`; its helpers win over `Load Stickers 2.gs`, same names, loads later).
- White band bug: `addStickerWithNumber` compared colour to "#FFFFFF" exactly, Product List gives "#ffffff" → white text on white. Now case-insensitive.
- Pack in barcode rounded to 2 decimals (`-0.83`, was `-0.8333…`).
- **Own barcode** (`Barcode 128.gs`): `code128Text_` = Code 128 (set C for digit runs ≥4, else B; checksum; stop) as characters for the Google font **Libre Barcode 128** (value 0 → Â, <95 → +32, ≥95 → +100). Checked: rendered with the real font and read back by zxing-cpp for 7 codes incl. "88008990-0.83", "8901719115219", "ABC-12". `ownBarcodeSetup_` finds the template cell (`#Ref_Novajet 24L Format` A1:F9) whose formula contains "barcode" (the mobiledemand =image(...)) + its merged size on Novajet 24L; `ownBarcodePut_` overwrites that cell on each sticker with the encoded text, font Libre Barcode 128, size fitted (0.92 × width / (0.33 em per symbol), ≤ 0.98 × height), centred, clipped. Switch off: `OWN_BARCODE_ = false`. Custom function `=BARCODE128(cell)` also available.
- **Fix (8 Oct 2026):** first try overwrote the LOGO cell (A1 of each sticker): its formula also contains "barcode". Now only a formula matching `symbology=` or `barcode/image` counts, and exactly one must exist in the template, else the website barcode is kept. Real barcode cell = template C5 (merged, e.g. O14:Q15 on the 3rd sticker of row 2). Logos come back on the next LOAD (template is copied each time).
- **Bigger barcode (8 Oct 2026):** first live try worked but was small (box-fit). Now font px = min(0.94 × width / em, height / 0.6) so the bars (0.59 em) fill the height; if the text box is taller than the cell it is top-aligned (only the empty descender part is cut). Barcode rows are forced back to their template heights (`setRowHeightsForced`, rows from the first sticker page) so a big font cannot stretch the sticker layout.
- **Barcode PICTURE instead of font (8 Oct 2026, pushed, not yet seen live):** the font route was unreliable (Sheets' vertical placement of big text cut the bars to a thin strip). Now `ownBarcodePut_` builds a PNG in the script (`code128Png_`: Code 128 module table `C128_PATTERNS_`, 1-bit grayscale, stored zlib + CRC32/Adler, 6 quiet modules each side, ~600+ px wide, height = width × cell h/w so it fills the cell) and puts it IN the cell via `SpreadsheetApp.newCellImage().setSourceUrl('data:image/png;base64,…')` (cached per code). Checked in node + zxing-cpp: 5 codes scan, also squashed. If the cell image throws → font fallback (box-fit size, middle). Row heights still forced.
