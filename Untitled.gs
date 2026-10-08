function importFromConversionSheet() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. Connect to your Inventory Spreadsheet
  const INVENTORY_SS_ID = '1h-Rm2RqJsG_k826d3flemOrYN73laSb5s7ngQn6-E3M'; // Your Inventory SS ID
  const invSS = SpreadsheetApp.openById(INVENTORY_SS_ID);
  const convSheet = invSS.getSheetByName("Conversion Sheet");
  
  if (!convSheet) {
    ui.alert("❌ Error: Could not find 'Conversion Sheet' in the Inventory file.");
    return;
  }
  
  // 2. Get the data from Conversion Sheet (Assuming Col B is Name, Col C is Pkts)
  const lastRow = convSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("⚠️ Conversion Sheet is empty.");
    return;
  }
  
  const data = convSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  
  // 3. Setup Print Console
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const work = ss.getSheetByName("Print Console");
  const list = ss.getSheetByName("Product List");
  
  let itemsProcessed = 0;
  
  // 4. Loop through each row in Conversion Sheet
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const productName = row[1]; // Column B
    const pkts = row[2];        // Column C
    
    // Skip empty rows
    if (!productName || !pkts) continue; 
    
    // Simulate human entry into B2 and C2
    work.getRange("B2").setValue(productName);
    work.getRange("C2").setValue(pkts);
    
    // Force Google Sheets to calculate the formulas in A2, D2, E2, F2, G2, etc.
    SpreadsheetApp.flush();
    Utilities.sleep(500); // Give formulas half a second to load
    
    // Run the silent insert
    const success = silentInsert_(work, list);
    if (success) itemsProcessed++;
  }
  
  // 5. Clean up B2 and C2 when done
  work.getRange("B2").clearContent();
  work.getRange("C2").clearContent();
  work.getRange("D2").setValue(1);
  work.getRange("F2").clearContent();
  work.getRange("G2").clearContent();
  
  ui.alert(`✅ Success! Imported ${itemsProcessed} items from the Conversion Sheet.`);
}

// ------------------------------------------------------------------
// SILENT INSERT LOGIC (Exact copy of insert3, but without UI alerts)
// ------------------------------------------------------------------
function silentInsert_(work, list) {
  const barcode = work.getRange("A2").getValue();
  const valueC2 = work.getRange("C2").getValue(); 
  const valueC5 = work.getRange("C5").getValue(); 
  const valueE2 = work.getRange("E2").getValue(); 
  const valueF2 = work.getRange("F2").getValue(); 
  const valueG2 = work.getRange("G2").getValue(); 
  const valueD2 = work.getRange("D2").getValue(); 
  
  if (!barcode || !valueC2) return false; 
  if ((valueC5 + valueC2) > 240) return false; // Exceeds limit

  const row = Math.max(work.getLastRow() + 1, 8);
  const prodlistData = list.getRange("B2:J3000").getValues();

  for (let i = 0; i < prodlistData.length; i++) {
    const productBarcode = prodlistData[i][0]; 

    if (productBarcode == barcode) {
      // Update Product List
      if (valueF2) list.getRange(i + 2, 7).setValue(valueF2); 
      if (valueG2) list.getRange(i + 2, 9).setValue(valueG2); 

      const latestH2 = work.getRange("H2").getValue();

      // Paste data down into row 8+
      work.getRange(row, 1).setValue(productBarcode); 
      work.getRange(row, 3).setValue(valueC2); 
      work.getRange(row, 4).setValue(valueD2); 
      work.getRange(row, 10).setValue(latestH2).setBackground(latestH2); 

      // Dates
      if (valueE2 && valueF2) {
        const pkd = new Date(valueE2); 
        work.getRange(row, 5).setValue(pkd).setNumberFormat('DD/MM/YYYY');
        work.getRange(row, 6).setValue(new Date(pkd.getTime() + (valueF2 * 86400000))); 
      }
      return true; 
    }
  }
  return false;
}