function insert() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var work = ss.getSheetByName("Print Console");
    var list = ss.getSheetByName("Product List");
    var barcode = work.getRange("A2").getValue();
    var row = work.getLastRow() + 1;
    var ui = SpreadsheetApp.getUi();

    if (work.getRange("B2").isBlank() || work.getRange("C2").isBlank()) {
        ui.alert('Kindly fill in Item Code and Sticker Qty');
        return;
    }

    if (work.getRange("A2").isBlank()) {
        ui.alert('Item code doesn\'t exist');
        return;
    }

    if ((work.getRange("C5").getValue() + work.getRange("C2").getValue()) >= 241) {
        ui.alert('Number of stickers exceeding the printing limit of 5 sheets or 120 stickers');
        return;
    }

    var prod = work.getRange("B2").getValue();
    var prodlist = list.getRange("A1:J3000").getValues();

    for (var i = 1; i < prodlist.length; i++) {
        if (prodlist[i][0] == prod) {
            var j = i + 1;
            list.getRange(j, 7).setValue(work.getRange("F2").getValue());
            list.getRange(j, 9).setValue(work.getRange("G2").getValue());
        }
    }

    if (work.getRange("F4").getValue() < 10) {
        for (var k = 0; k < prodlist.length; k++) {
            var vprodlist = prodlist[k];
            if (prodlist[k][1] == barcode) {
                work.getRange(row, 1).setValue(vprodlist[1]); // Item Code
                work.getRange(row, 2).setValue(vprodlist[2].toString().toUpperCase()); // Item Name
                work.getRange(row, 3).setValue(work.getRange("C2").getValue()); // Sticker
                work.getRange(row, 4).setValue(work.getRange("D2").getValue()); // Packing Size
                work.getRange(row, 10).setBackground(work.getRange("H2").getValue()); // Color code

                var pkd = new Date(work.getRange("E2").getValue());

                if (work.getRange("F2").getValue() > 0) {
                    work.getRange(row, 5).setValue(pkd).setNumberFormat('DD/MM/YYYY');
                    work.getRange(row, 6).setValue(new Date(pkd.getTime() + (work.getRange("F2").getValue() * 86400000)));
                }
            }
        }
    }

    work.getRange("C2").clearContent();
    work.getRange("D2").setValue(1);
    work.getRange("F2").clearContent();
    work.getRange("G2").clearContent();
    work.getRange("E2").setFormula("=K1");
    work.getRange('B2').activate();

    // Calculate the sum after inserting data
    var rangeC = work.getRange("C8:C100").getValues();
    var rangeH = work.getRange("H8:H100").getValues();

    var sum = 0;

    for (var i = 0; i < rangeC.length; i++) {
        sum += rangeC[i][0] * rangeH[i][0];
    }

    // Output the sum to cell H5
    work.getRange("H5").setValue(sum);
}
