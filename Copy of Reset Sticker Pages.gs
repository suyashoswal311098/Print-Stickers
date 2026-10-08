function Reset2x2() {
  var ss =SpreadsheetApp.getActiveSpreadsheet();
  var work =ss.getSheetByName("Novajet 24L");
  var rstk = ss.getSheetByName("#Ref_Novajet 24L Format")
  var rstk2 = ss.getSheetByName("formula of Novajet 24L")


  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("RESET STICKER PAGES", "Do you want to Reset the Sticker Pages?", ui.ButtonSet.YES_NO);
    
if (response==ui.Button.YES){



  //work.getRangeList(['A3:E3', 'A4:E4', 'B5:B7', 'A9:B9', 'D9:E9', 'C7:E7']).activate().clear({contentsOnly: true, skipFilteredRows: true});
  work.getRange("A1:R720").clear()

  //work.getRange(64,6).setValue(1);
  //work.getRange(64,12).setValue(1);
 // work.getRange(64,18).setValue(1);
  //work.getRange(136,6).setValue(2);
 //work.getRange(136,12).setValue(2);
  //work.getRange(136,18).setValue(2);
 // work.getRange(208,6).setValue(3);
 // work.getRange(208,12).setValue(3);
 // work.getRange(208,18).setValue(3);
 // work.getRange(280,6).setValue(4);
 // work.getRange(280,12).setValue(4);
 // work.getRange(280,18).setValue(4);
 // work.getRange(352,6).setValue(5);
 // work.getRange(352,12).setValue(5);
 // work.getRange(352,18).setValue(5);//
 


  ss.getSheetByName("Print Console").getRange('B2').activate();
  ss.getSheetByName("Print Console").getRange('B2').clearContent();

//Lengthy Work - Select and paste
  //work.getRange('A1:E360').activate();
  //work.getRange('A1:E9').copyTo(work.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  //work.getRange('G1:K360').activate();
  //work.getRange('A1:E9').copyTo(work.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  //work.getRange('M1:Q360').activate();
  //work.getRange('A1:E9').copyTo(work.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

}  
};