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
