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
