function UpdateNAME() {
  
  var ss =SpreadsheetApp.getActiveSpreadsheet();
  var imp =ss.getSheetByName("Import List");
  var prod =ss.getSheetByName("Product List");

  var tag = "Name Change";

  var iv = imp.getDataRange().getValues(); //import list values
  var pv = prod.getDataRange().getValues(); // product list values
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("UPDATE NAME", "Do you want to update NAMEs to the Product List?", ui.ButtonSet.YES_NO);
    
   if (response==ui.Button.YES){

     for (v=0;v<iv.length;v++){
       var ir = iv[v];

        if(ir[6]==tag){
          
          var icode = ir[0];
          var iname = ir[1];

          for (s=0;s<pv.length;s++){
            var pr = pv[s];

            if(pr[1]==icode){
               
            prod.getRange((s+1),3).setValue(iname);// s+1 as the product list contains one header row and 6 for no or MRP column
              
              }
          }
       
        }

     }


   }

}
