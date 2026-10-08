function AddProductsFromImportList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var imp = ss.getSheetByName("Import List");
  var prod = ss.getSheetByName("Product List");

  var importData = imp.getRange(4, 1, imp.getLastRow() - 3, imp.getLastColumn()).getValues(); // import list values
  var productData = prod.getDataRange().getValues(); // product list values

  // Get the column indexes based on the header row
  var headerRow = imp.getRange(1, 1, 1, imp.getLastColumn()).getValues()[0];
  var icodeIndex = headerRow.indexOf("icode");
  var inameIndex = headerRow.indexOf("iname");
  var brandIndex = headerRow.indexOf("brand");
  var nppIndex = headerRow.indexOf("npp");
  var mrpIndex = headerRow.indexOf("mrp");

  for (var v = 0; v < importData.length; v++) {
    var ir = importData[v];
    var icode = ir[icodeIndex];
    var iname = ir[inameIndex];
    var brand = ir[brandIndex];
    var npp = ir[nppIndex];
    var mrp = ir[mrpIndex];

    // Check if the product code already exists in the product list
    var productExists = false;
    for (var s = 0; s < productData.length; s++) {
      var pr = productData[s];
      if (pr[1] == icode) {
        productExists = true;
        break;
      }
    }

    // If the product code doesn't exist, add the new product
    if (!productExists) {
      var newRow = [icode, iname, brand, npp, mrp, "60", "", "White", "STORE IN DRY PLACE"];
      prod.appendRow(newRow);
      SpreadsheetApp.flush(); // Ensure the row is saved immediately
      Logger.log("New product added: " + iname);
    } else {
      Logger.log("Product with code " + icode + " already exists.");
    }
  }
}
