function SendToWorkLogCL() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName("CL");
  var sheet2 = ss.getSheetByName("Work Log");
  var repsheet = ss.getSheetByName("Representing");
  var rep = sheet1.getRange("D15");
  var repvalue = rep.getValue();
  var order = sheet1.getRange("D17");
  var ordervalue = order.getValue();
  
  var brand = sheet1.getRange("D16");
  var brandvalue = brand.getValue();
  
  if (ordervalue.indexOf('CNZ') === -1) {
  Browser.msgBox("Please enter a valid Order #");
  
  }else if (repvalue == 'Yes') {
  
  sheet1.getRange("A2:Y2").copyTo(repsheet.getRange(repsheet.getLastRow()+1,1,1,7), {contentsOnly:true});
  sheet1.getRange("A2:Y2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('A3:B39').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A3').setValue('Ctrl+SHIFT+V');
  SpreadsheetApp.getActiveSheet().getRange('D13').setValue('Yes');
  SpreadsheetApp.getActiveSheet().getRange('D14:D15').setValue('No');
  SpreadsheetApp.getActiveSheet().getRange('D17').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D18').setValue('0');
  SpreadsheetApp.getActiveSheet().getRange('D20').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('G12:G15').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('J12:J15').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D16').setFormula("=IF(A2=\"\",\"\",VLOOKUP(E18,'Data Validation'!C2:E204,3,FALSE))");

  }else{
  
  sheet1.getRange("A2:Y2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('A3:B39').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A3').setValue('Ctrl+SHIFT+V');
  SpreadsheetApp.getActiveSheet().getRange('D13').setValue('Yes');
  SpreadsheetApp.getActiveSheet().getRange('D14:D15').setValue('No');
  SpreadsheetApp.getActiveSheet().getRange('D17').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D18').setValue('0');
  SpreadsheetApp.getActiveSheet().getRange('D20').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D16').setFormula("=IF(A2=\"\",\"\",VLOOKUP(E18,'Data Validation'!C2:E204,3,FALSE))");

 }
 }
 
function SendToWorkLogAM() { //This is the script for the Amex sheet

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName("AM");
  var sheet2 = ss.getSheetByName("Work Log");
  var repsheet = ss.getSheetByName("Representing");
  var rep = sheet1.getRange("D15");
  var repvalue = rep.getValue();
  var order = sheet1.getRange("D17");
  var ordervalue = order.getValue();
  
  var brand = sheet1.getRange("D16");
  var brandvalue = brand.getValue();
  
  if (ordervalue.indexOf('CNZ') === -1) {
  Browser.msgBox("Please enter a valid Order #");
  
  }else if (repvalue == 'Yes') {
  
  sheet1.getRange("A2:Y2").copyTo(repsheet.getRange(repsheet.getLastRow()+1,1,1,7), {contentsOnly:true});
  sheet1.getRange("A2:Y2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('A3:B13').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A3').setValue('Ctrl+SHIFT+V');
  SpreadsheetApp.getActiveSheet().getRange('D13').setValue('Yes');
  SpreadsheetApp.getActiveSheet().getRange('D14:D15').setValue('No');
  SpreadsheetApp.getActiveSheet().getRange('D17').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D18').setValue('0');
  SpreadsheetApp.getActiveSheet().getRange('D20').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('G12:G15').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D16').setFormula("=IF(A2=\"\",\"\",VLOOKUP(E18,'Data Validation'!D2:E204,2,FALSE))");
  
  }else{
  
  sheet1.getRange("A2:Y2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('A3:B13').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A3').setValue('Ctrl+SHIFT+V');
  SpreadsheetApp.getActiveSheet().getRange('D13').setValue('Yes');
  SpreadsheetApp.getActiveSheet().getRange('D14:D15').setValue('No');
  SpreadsheetApp.getActiveSheet().getRange('D17').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D18').setValue('0');
  SpreadsheetApp.getActiveSheet().getRange('D20').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('G12:G15').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D16').setFormula("=IF(A2=\"\",\"\",VLOOKUP(E18,'Data Validation'!D2:E204,2,FALSE))");

 }
 }
 
function SendToWorkLogPP() { //This is the script for the PayPal sheet

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName("PP");
  var sheet2 = ss.getSheetByName("Work Log");
  var repsheet = ss.getSheetByName("Representing");
  var rep = sheet1.getRange("D15");
  var repvalue = rep.getValue();
  var order = sheet1.getRange("D17");
  var ordervalue = order.getValue();
  
  var brand = sheet1.getRange("D16");
  var brandvalue = brand.getValue();
  
  if (ordervalue.indexOf('CNZ') === -1) {
  Browser.msgBox("Please enter a valid Order #");
  
  }else if (repvalue == 'Yes') {
  
  sheet1.getRange("A2:Y2").copyTo(repsheet.getRange(repsheet.getLastRow()+1,1,1,7), {contentsOnly:true});
  sheet1.getRange("A2:Y2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('A3:A6').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A11:A23').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('B3:B23').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A3').setValue('Ctrl+SHIFT+V');
  SpreadsheetApp.getActiveSheet().getRange('D13').setValue('Yes');
  SpreadsheetApp.getActiveSheet().getRange('D14:D15').setValue('No');
  SpreadsheetApp.getActiveSheet().getRange('D17').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D18').setValue('0');
  SpreadsheetApp.getActiveSheet().getRange('D20').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('G12:G15').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A7').setValue('Amount:');
  SpreadsheetApp.getActiveSheet().getRange('A8').setValue('Trans Date:');
  SpreadsheetApp.getActiveSheet().getRange('A9').setValue('Case #:');
  SpreadsheetApp.getActiveSheet().getRange('A10').setValue('Reason Code:');
  
  }else{
  
  sheet1.getRange("A2:Y2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('A3:A6').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A11:A23').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('B3:B23').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A3').setValue('Ctrl+SHIFT+V');
  SpreadsheetApp.getActiveSheet().getRange('D13').setValue('Yes');
  SpreadsheetApp.getActiveSheet().getRange('D14:D15').setValue('No');
  SpreadsheetApp.getActiveSheet().getRange('D16').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D17').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D18').setValue('0');
  SpreadsheetApp.getActiveSheet().getRange('D20').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A7').setValue('Amount:');
  SpreadsheetApp.getActiveSheet().getRange('A8').setValue('Trans Date:');
  SpreadsheetApp.getActiveSheet().getRange('A9').setValue('Case #:');
  SpreadsheetApp.getActiveSheet().getRange('A10').setValue('Reason Code:');

 }
 }
 
function SendToWorkLogAD() { //This is the script for the Adyen sheet

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName("AD");
  var sheet2 = ss.getSheetByName("Work Log");
  var repsheet = ss.getSheetByName("Representing");
  var rep = sheet1.getRange("B21");
  var repvalue = rep.getValue();
  var order = sheet1.getRange("B23");
  var ordervalue = order.getValue();
  
  var brand = sheet1.getRange("B22");
  var brandvalue = brand.getValue();
  
  if (ordervalue.indexOf('CNZ') === -1) {
  Browser.msgBox("Please enter a valid Order #");
  
  }else if (repvalue == 'Yes') {
  
  sheet1.getRange("A2:Y2").copyTo(repsheet.getRange(repsheet.getLastRow()+1,1,1,7), {contentsOnly:true});
  sheet1.getRange("A2:Y2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('A3:F8').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A3').setValue('Ctrl+SHIFT+V');
  SpreadsheetApp.getActiveSheet().getRange('B18').setValue('Chargeback');
  SpreadsheetApp.getActiveSheet().getRange('B19').setValue('Yes');
  SpreadsheetApp.getActiveSheet().getRange('B20:B21').setValue('No');
  SpreadsheetApp.getActiveSheet().getRange('B22').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('B23').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('B24').setValue('0');
  SpreadsheetApp.getActiveSheet().getRange('B26').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('E18:E21').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('H18:H21').setValue('');
  
  }else{
  
  sheet1.getRange("A2:Y2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('A3:F8').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A3').setValue('Ctrl+SHIFT+V');
  SpreadsheetApp.getActiveSheet().getRange('B18').setValue('Chargeback');
  SpreadsheetApp.getActiveSheet().getRange('B19').setValue('Yes');
  SpreadsheetApp.getActiveSheet().getRange('B20:B21').setValue('No');
  SpreadsheetApp.getActiveSheet().getRange('B22').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('B23').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('B24').setValue('0');
  SpreadsheetApp.getActiveSheet().getRange('B26').setValue('');
  
 }
 }
