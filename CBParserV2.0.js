  var user = Session.getActiveUser().getUserLoginId();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet2 = ss.getSheetByName("Work Log");
  var repsheet = ss.getSheetByName("Representing");
  var updates = ss.getSheetByName("Updates");
  
function SendToWorkLogCL() { //This is the script for the ClientLine sheet

  var sheet1 = ss.getSheetByName("CL");
  var rep = sheet1.getRange("B20");
  var repvalue = rep.getValue();
  var order = sheet1.getRange("B16");
  var ordervalue = order.getValue();
  var brand = sheet1.getRange("B11");
  var purchase = sheet1.getRange("B15");
  var stage = sheet1.getRange("B9");
  
  SpreadsheetApp.getActiveSheet().getRange('B27').setValue(user);
  
  if (ordervalue.indexOf('CNZ') === -1) {
  Browser.msgBox("Please enter a valid Order #");
  return
  }else if (brand.getValue() =="") {
  Browser.msgBox("Please make sure the 'Brand' field is populated");
  return
  }else if (purchase.getValue() =="") {
  Browser.msgBox("Please make sure the 'Purchase Amount' field is populated");
  return
  }
  if (stage.getValue() == 'Graduated Chargeback' || stage.getValue() == 'Second Chargeback' ||  stage.getValue() == 'Pre-Arbitration') {
  sheet1.getRange("A2:X2").copyTo(updates.getRange(updates.getLastRow()+1,1,1,7), {contentsOnly:true});
  }
  if (repvalue == 'Yes') {
  sheet1.getRange("A2:X2").copyTo(repsheet.getRange(repsheet.getLastRow()+1,1,1,7), {contentsOnly:true});
  }
  if (stage.getValue() == 'Retrieval' || stage.getValue() == 'Chargeback') {
  sheet1.getRange("A2:X2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  }
  SpreadsheetApp.getActiveSheet().getRange('B11').setFormula("=IF(A2=\"\",\"\",VLOOKUP($D$18,'Data Validation'!$C$2:$E$204,3,FALSE))"); //Brand
  SpreadsheetApp.getActiveSheet().getRange('B9').setFormula("=IF($A$2=\"\",\"\",IF(RegExMatch($E$26,\"First Chargeback\"),\"Chargeback\",IF(RegExMatch($E$26,\"Second Chargeback\"),\"Second Chargeback\",IF(RegExMatch($E$26,\"Pre-Arbitration\"),\"Pre-Arbitration\",IF(RegExMatch($E$27,\"Retrieval\"),\"Retrieval\",\"\")))))"); //Stage Reached
  SpreadsheetApp.getActiveSheet().getRange('B16').setValue(''); //Order #
  SpreadsheetApp.getActiveSheet().getRange('B15').setFormula("=IF($B$14=\"USD\",SUBSTITUTE(SUBSTITUTE($B$5,\"USD\",\"\"),\" \",\"\"))"); //Purchase Amount
  SpreadsheetApp.getActiveSheet().getRange('B18').setValue('Yes'); //Auto'd
  SpreadsheetApp.getActiveSheet().getRange('B19:B20').setValue('No'); //Verified | Represented
  SpreadsheetApp.getActiveSheet().getRange('B21:B26').setValue(''); //Rep Reason --> Documentation 4
  SpreadsheetApp.getActiveSheet().getRange('B17').setValue('0'); //Balance
  SpreadsheetApp.getActiveSheet().getRange('E3:F39').setValue(''); //Paste Section
  SpreadsheetApp.getActiveSheet().getRange('E3').setValue('Ctrl+SHIFT+V'); //Paste Cell
  }
 
function SendToWorkLogAM() { //This is the script for the Amex sheet

  var sheet1 = ss.getSheetByName("AM");
  var rep = sheet1.getRange("B20");
  var repvalue = rep.getValue();
  var order = sheet1.getRange("B16");
  var ordervalue = order.getValue();
  var brand = sheet1.getRange("B11");
  var purchase = sheet1.getRange("B15");
  
  SpreadsheetApp.getActiveSheet().getRange('B27').setValue(user);
  
  if (ordervalue.indexOf('CNZ') === -1) {
  Browser.msgBox("Please enter a valid Order #");
  return
  }else if (brand.getValue() =="") {
  Browser.msgBox("Please make sure the 'Brand' field is populated");
  return
  }else if (purchase.getValue() =="") {
  Browser.msgBox("Please make sure the 'Purchase Amount' field is populated");
  return
  }
  
  if (repvalue == 'Yes') {
  
  sheet1.getRange("A2:X2").copyTo(repsheet.getRange(repsheet.getLastRow()+1,1,1,7), {contentsOnly:true});
  
  }
  
  sheet1.getRange("A2:X2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('B11').setFormula("=IF(A2=\"\",\"\",VLOOKUP(D18,'Data Validation'!D2:E204,2,FALSE))"); //Brand
  SpreadsheetApp.getActiveSheet().getRange('B9').setFormula("=IF($A$2=\"\",\"\",IF($E$13=\"IQ\",\"Retrieval\",\"Chargeback\"))"); //Stage Reached
  SpreadsheetApp.getActiveSheet().getRange('B16').setValue(''); //Order #
  SpreadsheetApp.getActiveSheet().getRange('B14').setValue('USD'); //Currency
  SpreadsheetApp.getActiveSheet().getRange('B18').setValue('Yes'); //Auto'd
  SpreadsheetApp.getActiveSheet().getRange('B19:B20').setValue('No'); //Verification | Represented
  SpreadsheetApp.getActiveSheet().getRange('B21:B26').setValue(''); //Rep Reason --> Documentation 4
  SpreadsheetApp.getActiveSheet().getRange('B17').setValue('0'); //Balance
  SpreadsheetApp.getActiveSheet().getRange('E3:F13').setValue(''); //Paste area
  SpreadsheetApp.getActiveSheet().getRange('E3').setValue('Ctrl+SHIFT+V'); //Paste cell

  }
 
function SendToWorkLogPP() { //This is the script for the PayPal sheet

  var sheet1 = ss.getSheetByName("PP");
  var rep = sheet1.getRange("B20");
  var repvalue = rep.getValue();
  var order = sheet1.getRange("B16");
  var ordervalue = order.getValue();
  var brand = sheet1.getRange("B15");
  var purchase = sheet1.getRange("B14");
  
  SpreadsheetApp.getActiveSheet().getRange('B27').setValue(user);
  
  if (ordervalue.indexOf('CNZ') === -1) {
  Browser.msgBox("Please enter a valid Order #");
  return
  }else if (brand.getValue() =="") {
  Browser.msgBox("Please make sure the 'Brand' field is populated");
  return
  }else if (purchase.getValue() =="") {
  Browser.msgBox("Please make sure the 'Purchase Amount' field is populated");
  return
  }
  
  if (repvalue == 'Yes') {
  
  sheet1.getRange("A2:X2").copyTo(repsheet.getRange(repsheet.getLastRow()+1,1,1,7), {contentsOnly:true});
  
  }
  
  sheet1.getRange("A2:X2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('B15').setValue(''); //Brand
  SpreadsheetApp.getActiveSheet().getRange('B16').setValue(''); //Order #
  SpreadsheetApp.getActiveSheet().getRange('B13').setValue('USD'); //Currency
  SpreadsheetApp.getActiveSheet().getRange('B18').setValue('Yes'); //Auto'd
  SpreadsheetApp.getActiveSheet().getRange('B19:B20').setValue('No'); //Verification | Represented
  SpreadsheetApp.getActiveSheet().getRange('B21:B26').setValue('');  //Rep Reason --> Documentation 4
  SpreadsheetApp.getActiveSheet().getRange('B17').setValue('0'); //Balance
  SpreadsheetApp.getActiveSheet().getRange('E3:F14').setValue(''); // Paste Area
  SpreadsheetApp.getActiveSheet().getRange('E3').setValue('Ctrl+SHIFT+V'); //Paste Cell
  SpreadsheetApp.getActiveSheet().getRange('E7').setValue('Amount:');
  SpreadsheetApp.getActiveSheet().getRange('E8').setValue('Trans Date:');
  SpreadsheetApp.getActiveSheet().getRange('E9').setValue('Case #:');
  SpreadsheetApp.getActiveSheet().getRange('E10').setValue('Reason Code:');
  SpreadsheetApp.getActiveSheet().getRange('E11').setValue('Dispute Date:');
  
  }
 
function SendToWorkLogAD() { //This is the script for the Adyen sheet

  var sheet1 = ss.getSheetByName("AD");
  var rep = sheet1.getRange("B20");
  var repvalue = rep.getValue();
  var order = sheet1.getRange("B16");
  var ordervalue = order.getValue();
  var brand = sheet1.getRange("B15");
  var purchase = sheet1.getRange("B14");
  
  SpreadsheetApp.getActiveSheet().getRange('B27').setValue(user);
  
  if (ordervalue.indexOf('CNZ') === -1) {
  Browser.msgBox("Please enter a valid Order #");
  return
  }else if (brand.getValue() =="") {
  Browser.msgBox("Please make sure the 'Brand' field is populated");
  return
  }else if (purchase.getValue() =="") {
  Browser.msgBox("Please make sure the 'Purchase Amount' field is populated");
  return
  }
  
  if (repvalue == 'Yes') {
  
  sheet1.getRange("A2:Y2").copyTo(repsheet.getRange(repsheet.getLastRow()+1,1,1,7), {contentsOnly:true});
  
  }
  
  sheet1.getRange("A2:X2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('B15').setValue(''); //Brand
  SpreadsheetApp.getActiveSheet().getRange('B9').setValue('Chargeback'); //Stage Reached
  SpreadsheetApp.getActiveSheet().getRange('B16').setValue(''); //Order #
  SpreadsheetApp.getActiveSheet().getRange('B18').setValue('Yes'); //Auto'd
  SpreadsheetApp.getActiveSheet().getRange('B19:B20').setValue('No'); //Verification | Represented
  SpreadsheetApp.getActiveSheet().getRange('B21:B26').setValue(''); //Rep Reason --> Documentation 4
  SpreadsheetApp.getActiveSheet().getRange('B17').setValue('0'); //Balance
  SpreadsheetApp.getActiveSheet().getRange('E3:I8').setValue(''); //Paste area
  SpreadsheetApp.getActiveSheet().getRange('E3').setValue('Ctrl+SHIFT+V'); //Paste cell
  
 }
 
 /**
 * Return a 0-based array index corresponding to a spreadsheet column
 * label, as in A1 notation.
 *
 * @param {String}    colA1    Column label to be converted.
 *
 * @return {Number}            0-based array index.
 */
function ColA1ToIndex( colA1 ) {
  if (typeof colA1 !== 'string' || colA1.length > 2) 
    throw new Error( "Expected column label." );

  var A = "A".charCodeAt(0);

  var number = colA1.charCodeAt(colA1.length-1) - A;
  if (colA1.length == 2) {
    number += 26 * (colA1.charCodeAt(0) - A + 1);
  }
  return number;
}

// ... TEST CODE for more efficient value setting

/**
 * Return a 0-based array index corresponding to a spreadsheet row
 * number, as in A1 notation.
 *
 * @param {Number}    rowA1    Row number to be converted.
 *
 * @return {Number}            0-based array index.
 */
function RowA1ToIndex( rowA1 ) {
  return rowA1 - 1;
}

function SendToWorkLogCLTEST() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName("CL");
  var sheet2 = ss.getSheetByName("Work Log");
  var repsheet = ss.getSheetByName("Representing");
  var rep = sheet1.getRange("D18");
  var repvalue = rep.getValue();

  var clRange = sheet1.getRange("A2:X2");
  if (repvalue == 'Yes') {
    clRange.copyTo(repsheet.getRange(repsheet.getLastRow() + 1, 1, 1, 7), {
      contentsOnly: true
    });
  }

  // This part does not need to be in an if/then/else, because it's always done.
  clRange.copyTo(sheet2.getRange(sheet2.getLastRow() + 1, 1, 1, 7), {
    contentsOnly: true
  });

  // Only need this block once, instead of two identical copies.
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  values[3-1][ColA1ToIndex('A')] = 'Ctrl+SHIFT+V';

  // ...getRange('A3:B37').setValue('') handled in loops
  for (var col=ColA1ToIndex('A'); col <= ColA1ToIndex('B'); col++) {
    for (var row=(3-1); row<=(37-1); row++) {
      values[row][col] = '';
    }
  }
  values[13-1][ColA1ToIndex('D')] = '';
  values[14-1][ColA1ToIndex('D')] = '';
  values[15-1][ColA1ToIndex('D')] = '';
  values[16-1][ColA1ToIndex('D')] = 'Yes';
  values[17-1][ColA1ToIndex('D')] = 'No';
  values[18-1][ColA1ToIndex('D')] = 'No';
  values[19-1][ColA1ToIndex('D')] = '';
  values[20-1][ColA1ToIndex('D')] = '';
  values[21-1][ColA1ToIndex('D')] = '';
  values[22-1][ColA1ToIndex('D')] = '';
  values[23-1][ColA1ToIndex('D')] = '';
  values[24-1][ColA1ToIndex('D')] = '';
  values[25-1][ColA1ToIndex('D')] = '0';

  // Finally, one service call to write ALL values. Fast!
  dataRange.setValues(values);
  // Formulas would have been overwritten by values, so need to be refreshed
  sheet.getRange('D11').setFormula("=IF(A2=\"\",\"\",VLOOKUP($D$29,'Data Validation'!C2:E204,3,FALSE))");
}
