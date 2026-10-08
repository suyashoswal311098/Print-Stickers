/**
 * This script automates the generation and printing of stickers based on the data in the "Print Console" sheet.
 *
 * Features:
 * 1. Ensures sufficient stickers are available and prompts for additional stickers if needed.
 * 2. Reads product details such as item code, name, packing size, MRP, NPP, brand, packaging date, expiry date, and notes.
 * 3. Updates the "Product List" sheet with the latest MRP, NPP, and item name for each product.
 * 4. Populates the "Novajet 24L" sheet with stickers for each product:
 *    - Aligns stickers in three columns per row.
 *    - Adds color coding for the product name and background.
 *    - Displays item details like MRP, NPP, packaging date, and expiry date.
 * 5. Creates a separator between stickers for different products.
 * 6. Logs details of each printed sticker in the "Print Log" sheet, including timestamps and user information.
 * 
 * This script is designed for organized and efficient sticker printing for multiple products in bulk.
 */

function load() {
  
  var ss =SpreadsheetApp.getActiveSpreadsheet();
  var work =ss.getSheetByName("Print Console");
  var list =ss.getSheetByName("Product List");
  var stkr =ss.getSheetByName("Novajet 24L");
  var log =ss.getSheetByName ("Print Log");
  var rl =log.getLastRow()+1;

  var rwt =work.getLastRow(); //total row of data in Print Console

  var stc =1; //sticker column 1 or 2 or 0
  var sn =0; 

  var ic =7; //item code row no in 1st stkr
  var iname =4; // item name row no in 1st stkr
  var ib =3; // item barcode row no in 1st stkr
  var imrp =5; //MRP row no in 1st stkr
  var inpp = 6; //NPP row no in 1st stkr
  var idt =9; //Pkg Date row no in 1st stkr
  var inote =8; //Item Note row no in 1st stkr

  var iseperator = 1; //item seperator


  var ui = SpreadsheetApp.getUi();

    if (work.getRange("C4").getValue()>0){
      var response = ui.alert(work.getRange("C4").getValue() + " MORE STICKERS NEEDED","Kindly add " + work.getRange("C4").getValue() + " more stickers, to complete the page.",ui.ButtonSet.OK_CANCEL);

      if (response==ui.Button.OK){
        return;
      }
    }

var response = ui.alert("LOAD STICKERS", "Do you want to Load the Stickers?", ui.ButtonSet.YES_NO);
    
if (response==ui.Button.YES){
 
 for (var rw=8; rw<=rwt; rw++){

    var sn= work.getRange(rw,3).getValue(); //no of stickers of a particular product

    var wcode = work.getRange(rw,1).getValue(); //Item code
    var wname = work.getRange(rw,2).getValue(); //Item Name
    var wpack = work.getRange(rw,4).getValue();//Packing size
    var wbrand = work.getRange(rw,7).getValue(); //Brand
    var wmrp = work.getRange(rw,9).getValue()*wpack; //MRP
    var wnpp = work.getRange(rw,8).getValue()*wpack; //NPP
    var wnote = work.getRange(rw,11).getValue(); //Item Note

    var clr = work.getRange(rw,10).getValue(); //colour code

    var pk = new Date(work.getRange(rw,5).getValue()); //packaging
    var ex = new Date(work.getRange(rw,6).getValue()); //expiry

    //Update MRP, NPP and Brand
var prodlist =list.getRange("A1:J3000").getValues();
     
      for (var i=0; i<prodlist.length; i++){

       if (prodlist[i][1]==wcode){
          var j=i+1;
          list.getRange(j,5).setValue(wnpp); //NPP Update
          list.getRange(j,6).setValue(wmrp); //MRP Update
          list.getRange(j,3).setValue(wname.toString().toLowerCase);//Item Name
        }
      }

    for (var k=1; k<=sn; k++){
       
     //for the sticker in column 1 
      if (stc==1){
        
  if(wpack!=1){
     stkr.getRange(ic,3).setValue(wcode+"-"+wpack) ; //barcode
  }
  else if(wpack=1){
 stkr.getRange(ic,3).setValue(wcode)
  }

        if (clr=='#FFFFFF'){
          stkr.getRange(iname,1).setValue(wname).setFontColor('#000000'); //Item Name
        }
        else if (clr!='#FFFFFF'){
          stkr.getRange(iname,1).setValue(wname).setBackground(clr).setFontColor('#ffffff'); //Item Name
        }

        stkr.getRange(ib,1).setValue(wbrand); //Item Brand
        stkr.getRange(inote,1).setValue(wnote); //NPP
        stkr.getRange(imrp,2).setValue("₹"+wmrp+"/-"); //MRP
        stkr.getRange(inpp,2).setValue("₹"+wnpp+"/-"); //NPP
        
        if(work.getRange(rw,6).isBlank()==false){


          stkr.getRange(idt,1).setValue("PKD: "+pk.getDate()+"/"+(pk.getMonth()+1)+"/"+pk.getFullYear()); //pkg dt
          stkr.getRange(idt,4).setValue("EXP: "+ex.getDate()+"/"+(ex.getMonth()+1)+"/"+ex.getFullYear()); //exp dt
        }
       //code to deploy particular colour

        stc = 2;
      }

     //for the sticker in column 2
      else if (stc==2){

  if(wpack!=1){
   stkr.getRange(ic,9).setValue(wcode+"-"+wpack) ; //barcode
  }
  else if(wpack=1){
  stkr.getRange(ic,9).setValue(wcode)
  }
        if (clr=='#FFFFFF'){
          stkr.getRange(iname,7).setValue(wname).setFontColor('#000000'); //Item Name
        }
        else if (clr!='#FFFFFF'){
          stkr.getRange(iname,7).setValue(wname).setBackground(clr).setFontColor('#ffffff'); //Item Name
        }

        stkr.getRange(ib,7).setValue(wbrand); //Item Brand
        stkr.getRange(inote,7).setValue(wnote); //Item Note
        stkr.getRange(imrp,8).setValue("₹"+wmrp+"/-"); //MRP
        stkr.getRange(inpp,8).setValue("₹"+wnpp+"/-"); //NPP
      
        if(work.getRange(rw,6).isBlank()==false){

          stkr.getRange(idt,7).setValue("PKD: "+  pk.getDate()+"/"+(pk.getMonth()+1)+"/"+pk.getFullYear()); //pkg dt
          stkr.getRange(idt,10).setValue("EXP: "+ex.getDate()+"/"+(ex.getMonth()+1)+"/"+ex.getFullYear()); //exp dt
        }
       //code to deploy particular colour

        stc = 3;
      }

     //for the sticker in column 3
      else if (stc==3){

        
   if(wpack!=1){
   stkr.getRange(ic,15).setValue(wcode+"-"+wpack) ; //barcode
  }
  else if(wpack=1){
  stkr.getRange(ic,15).setValue(wcode)
  }

   if (clr=='#FFFFFF'){
          stkr.getRange(iname,13).setValue(wname).setFontColor('#000000'); //Item Name
        }
        else if (clr!='#FFFFFF'){
          stkr.getRange(iname,13).setValue(wname).setBackground(clr).setFontColor('#ffffff'); //Item Name
        }

       
        
        stkr.getRange(ib,13).setValue(wbrand); //Item Brand
        stkr.getRange(inote,13).setValue(wnote); //Item Note
        stkr.getRange(imrp,14).setValue("₹"+wmrp+"/-"); //MRP
        stkr.getRange(inpp,14).setValue("₹"+wnpp+"/-"); //NPP
      
        if(work.getRange(rw,6).isBlank()==false){

          stkr.getRange(idt,13).setValue("PKD: "+pk.getDate()+"/"+(pk.getMonth()+1)+"/"+pk.getFullYear()); //pkg dt
          stkr.getRange(idt,16).setValue("EXP: "+ex.getDate()+"/"+(ex.getMonth()+1)+"/"+ex.getFullYear()); //exp dt
        }
       //code to deploy particular colour

        stc = 1;

        ic = ic+9;
        iname = iname+9;
        ib = ib+9;
        imrp = imrp+9;
        inpp = inpp+9;
        idt = idt+9;
        inote = inote+9;
        iseperator = iseperator+9;
      }
    
    }

   //Sticker Seperator
      if (stc==2){
        stkr.getRange(iseperator,6).setValue("Item No "+(rw-7)).setBackground("#000000").setFontColor('#ffffff');
      }
      else if (stc==3){
        stkr.getRange(iseperator,12).setValue("Item No "+(rw-7)).setBackground("#000000").setFontColor('#ffffff');
      }


   //Create the sticker print record
      log.getRange(rl,1).setValue(wcode); //Item code
      log.getRange(rl,2).setValue(wname); //Item Name
      log.getRange(rl,3).setValue(sn); //Stkr Qty
      log.getRange(rl,4).setValue(work.getRange(rw,4).getValue()); //Packaging Size
      log.getRange(rl,5).setValue(pk.getDate()+"/"+(pk.getMonth()+1)+"/"+pk.getFullYear()); //pkg dt
      log.getRange(rl,6).setValue(ex.getDate()+"/"+(ex.getMonth()+1)+"/"+ex.getFullYear()); //exp dt
      log.getRange(rl,7).setValue(wbrand); //Brand
      log.getRange(rl,9).setValue(wmrp); //MRP
      log.getRange(rl,8).setValue(wnpp); //NPP
      log.getRange(rl,10).setValue(clr); //ColourCode

      log.getRange(rl,11).setValue(new Date()).setNumberFormat('dd-mm-yyyy hh:mm'); //Time Stamp
      log.getRange(rl,12).setValue(Session.getActiveUser().getEmail());  //item code

    rl++;
  }

  
}
}
