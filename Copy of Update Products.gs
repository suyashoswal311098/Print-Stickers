// function updateProductList2() {
//   const ss           = SpreadsheetApp.getActiveSpreadsheet();
//   const ui           = SpreadsheetApp.getUi();
//   const importSheet  = ss.getSheetByName('Import List');
//   const productSheet = ss.getSheetByName('Product List');
 
//   if (!importSheet || !productSheet) {
//     ui.alert("Make sure both sheets exist.");
//     return;
//   }

//   const HEADER_ROWS     = 3;
//   const FIRST_ROW       = HEADER_ROWS + 1;
//   const NUM_IMPORT_COLS = 9;    
//   const totalRows       = importSheet.getLastRow();
//   const dataRows        = totalRows - HEADER_ROWS;
//   if (dataRows <= 0) {
//     ui.alert("No rows under header.");
//     return;
//   }
//   const rawData = importSheet.getRange(FIRST_ROW, 1, dataRows, NUM_IMPORT_COLS).getValues();

//   const lastProdRow = productSheet.getLastRow();
//   const prodCols    = Math.max(productSheet.getLastColumn(), 9); 
//   const prodRows    = Math.max(0, lastProdRow - 1);
//   const prodData    = prodRows ? productSheet.getRange(2, 1, prodRows, prodCols).getValues() : [];

//   const prodMap = new Map();
//   prodData.forEach((r, i) => {
//     if (r[1]) prodMap.set(r[1], i);
//     if (r[2]) prodMap.set(r[2], i);
//   });

//   const toAppend = [];
//   const updated  = [];

//   const COL_IN  = { bar:0, name:1, brand:2, npp:3, mrp:4, shelf:5, rating:6, action:8 };
//   const COL_OUT = { bar:1, name:2, brand:3, npp:4, mrp:5, shelf:6, rating:8 };

//   rawData.forEach(rowIn => {
//     const bar = rowIn[COL_IN.bar];
//     const name = rowIn[COL_IN.name];
//     if (!bar && !name) return; 

//     const rawAct = (rowIn[COL_IN.action]||"").trim();
//     if (!rawAct) return;                       

//     const flags = rawAct.split(',').map(f => f.trim().toUpperCase());
//     const isNew = flags.includes('NEW PRODUCT');
//     const idx   = prodMap.get(bar || name);

//     if (isNew) {
//       const newRow = Array(prodCols).fill("");
//       newRow[COL_OUT.bar]    = bar;
//       newRow[COL_OUT.name]   = name;
//       newRow[COL_OUT.brand]  = rowIn[COL_IN.brand];
//       newRow[COL_OUT.npp]    = rowIn[COL_IN.npp];
//       newRow[COL_OUT.mrp]    = rowIn[COL_IN.mrp];
//       newRow[COL_OUT.shelf]  = rowIn[COL_IN.shelf];
//       newRow[COL_OUT.rating] = rowIn[COL_IN.rating]; 
//       toAppend.push(newRow);
//       return;
//     }

//     if (idx != null) {
//       const out = prodData[idx];
//       let hasChanged = false;
//       if (flags.includes('NP CHANGE')) { out[COL_OUT.npp] = rowIn[COL_IN.npp]; hasChanged = true; }
//       if (flags.includes('MRP CHANGE')) { out[COL_OUT.mrp] = rowIn[COL_IN.mrp]; hasChanged = true; }
//       if (flags.includes('RATING CHANGE')) { out[COL_OUT.rating] = rowIn[COL_IN.rating]; hasChanged = true; }
//       if (hasChanged) updated.push(bar||name);
//     }
//   });

//   // --- WRITE UPDATES (Protecting Column H) ---
//   if (prodData.length) {
//     // 1. Write B through G (Indices 1 to 6)
//     const blockBG = prodData.map(r => r.slice(1, 7)); 
//     productSheet.getRange(2, 2, blockBG.length, 6).setValues(blockBG);

//     // 2. Write Column I only (Index 8)
//     const blockI = prodData.map(r => [r[8]]);
//     productSheet.getRange(2, 9, blockI.length, 1).setValues(blockI);
//   }

//   // --- APPEND NEW ---
//   if (toAppend.length) {
//     // For new rows, we append B-G and I. 
//     // Column H will stay empty (or you can let your formula drag down naturally)
//     const appendBG = toAppend.map(r => r.slice(1, 7));
//     productSheet.getRange(lastProdRow + 1, 2, appendBG.length, 6).setValues(appendBG);

//     const appendI = toAppend.map(r => [r[8]]);
//     productSheet.getRange(lastProdRow + 1, 9, appendI.length, 1).setValues(appendI);
//   }

//   ui.alert(`Done. Updated: ${updated.length}, Added: ${toAppend.length}`);
// }