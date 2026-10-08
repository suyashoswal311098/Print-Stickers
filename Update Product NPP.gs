  function UpdateNPP() {
  
  var ss =SpreadsheetApp.getActiveSpreadsheet();
  var imp =ss.getSheetByName("Import List");
  var prod =ss.getSheetByName("Product List");

  var tag = "NPP Change";
  
  var iv = imp.getDataRange().getValues();
  var pv = prod.getDataRange().getValues();
    
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("UPDATE NPP", "Do you want to update NPP to the Product List?", ui.ButtonSet.YES_NO);
    
   if (response==ui.Button.YES){

     for (v=0;v<iv.length;v++){
       var ir = iv[v];

        if(ir[8]==tag){
          
          var icode = ir[0];
          var inpp = ir[3];

          for (s=0;s<pv.length;s++){
            var pr = pv[s];

            if(pr[1]==icode){
               
            prod.getRange((s+1),5).setValue(inpp);// s+1 as the product list contains one header row and 6 for no or MRP column
              
              }
          }
       
        }

     }


   }

}
