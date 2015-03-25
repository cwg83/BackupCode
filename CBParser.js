# CBParserScript
Google docs javascript for the chargeback and logging parser

function onEdit(event)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = event.source.getActiveSheet().getSheetName();
  
   if (sheetName.match(/CL/) == null)
      // These aren't the droids you're looking for...
      return;
      
  var sheetCL = ss.getSheetByName("CL");
  var CellRow = SpreadsheetApp.getActiveRange().getRow();
  var CellColumn = SpreadsheetApp.getActiveRange().getColumn();
  var CLD6 = "=A21";
  var CLD7 = "=IF(A2=\"\",\"\",SPLIT(A20,\" \"))";
  var CLD8 = "=IF(A2=\"\",\"\",A17-1)";
  var CLD9 = "=IF(A2=\"\",\"\",A14)";
  var CLD10 = "=IF(A2=\"\",\"\",SPLIT(A5,\" \"))";
  var CLD11 = "=A8";  
  
  if (CellColumn == 4 & CellRow == 6){
    sheetCL.getRange(CellRow, CellColumn).setFormula(CLD6);
    Browser.msgBox("BAD HUMAN! NO CHANGE FORMULA!");
  }
  if (CellColumn == 4 & CellRow == 7){
    sheetCL.getRange(CellRow, CellColumn).setFormula(CLD7);
    Browser.msgBox("BAD HUMAN! NO CHANGE FORMULA!");
  }
  if (CellColumn == 4 & CellRow == 8){
    sheetCL.getRange(CellRow, CellColumn).setFormula(CLD8);
    Browser.msgBox("BAD HUMAN! NO CHANGE FORMULA!");
  }
  if (CellColumn == 4 & CellRow == 9){
    sheetCL.getRange(CellRow, CellColumn).setFormula(CLD9);
    Browser.msgBox("BAD HUMAN! NO CHANGE FORMULA!");
  }
  if (CellColumn == 4 & CellRow == 10){
    sheetCL.getRange(CellRow, CellColumn).setFormula(CLD10);
    Browser.msgBox("BAD HUMAN! NO CHANGE FORMULA!");
  }  
  if (CellColumn == 4 & CellRow == 11){
    sheetCL.getRange(CellRow, CellColumn).setFormula(CLD11);
    Browser.msgBox("BAD HUMAN! NO CHANGE FORMULA!");
  }  
}

function SendToWorkLogCL() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName("CL");
  var sheet2 = ss.getSheetByName("Work Log");
  var repsheet = ss.getSheetByName("Representing");
  var rep = sheet1.getRange("D15");
  var repvalue = rep.getValue();
  
  if (ordervalue == "") {
  Browser.msgBox("Please enter a valid Order #");
  
  }else if (repvalue == 'Yes') {
  
  sheet1.getRange("A2:W2").copyTo(repsheet.getRange(repsheet.getLastRow()+1,1,1,7), {contentsOnly:true});
  sheet1.getRange("A2:W2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('A3:B23').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A3').setValue('Ctrl+SHIFT+V');
  SpreadsheetApp.getActiveSheet().getRange('D12').setValue('Chargeback');
  SpreadsheetApp.getActiveSheet().getRange('D13').setValue('Yes');
  SpreadsheetApp.getActiveSheet().getRange('D14:D15').setValue('No');
  SpreadsheetApp.getActiveSheet().getRange('D17').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D18').setValue('0');
  SpreadsheetApp.getActiveSheet().getRange('D20').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('G12:G15').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D16').setFormula("=IF(A2=\"\",\"\",VLOOKUP(E18,'Data Validation'!C2:E204,3,FALSE))");

  }else{
  
  sheet1.getRange("A2:W2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('A3:B23').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A3').setValue('Ctrl+SHIFT+V');
  SpreadsheetApp.getActiveSheet().getRange('D12').setValue('Chargeback');
  SpreadsheetApp.getActiveSheet().getRange('D13').setValue('Yes');
  SpreadsheetApp.getActiveSheet().getRange('D14:D15').setValue('No');
  SpreadsheetApp.getActiveSheet().getRange('D17').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D18').setValue('0');
  SpreadsheetApp.getActiveSheet().getRange('D20').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D16').setFormula("=IF(A2=\"\",\"\",VLOOKUP(E18,'Data Validation'!C2:E204,3,FALSE))");
  
 }
 }
 
function SendToWorkLogAM() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName("AM");
  var sheet2 = ss.getSheetByName("Work Log");
   var repsheet = ss.getSheetByName("Representing");
  var rep = sheet1.getRange("D16");
  var repvalue = rep.getValue();
  
  if (repvalue == 'Yes') {
  
  sheet1.getRange("A2:W2").copyTo(repsheet.getRange(repsheet.getLastRow()+1,1,1,7), {contentsOnly:true});
  sheet1.getRange("A2:W2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('A3:B13').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A3').setValue('Ctrl+V');
  SpreadsheetApp.getActiveSheet().getRange('D13').setValue('Yes');
  SpreadsheetApp.getActiveSheet().getRange('D14:D15').setValue('No');
  SpreadsheetApp.getActiveSheet().getRange('D17').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D18').setValue('0');
  SpreadsheetApp.getActiveSheet().getRange('D20').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('G12:G15').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D16').setFormula("=IF(A2=\"\",\"\",VLOOKUP(E18,'Data Validation'!D2:E204,2,FALSE))");
  
  }else{
  
  
  sheet1.getRange("A2:W2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('A3:B13').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A3').setValue('Ctrl+V');
  SpreadsheetApp.getActiveSheet().getRange('D13').setValue('Yes');
  SpreadsheetApp.getActiveSheet().getRange('D14:D15').setValue('No');
  SpreadsheetApp.getActiveSheet().getRange('D17').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D18').setValue('0');
  SpreadsheetApp.getActiveSheet().getRange('D20').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('G12:G15').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D16').setFormula("=IF(A2=\"\",\"\",VLOOKUP(E18,'Data Validation'!D2:E204,2,FALSE))");
 }
 }
 
function SendToWorkLogPP() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName("PP");
  var sheet2 = ss.getSheetByName("Work Log");
  var repsheet = ss.getSheetByName("Representing");
  var rep = sheet1.getRange("D15");
  var repvalue = rep.getValue();
  
  if (repvalue == 'Yes') {
  
  sheet1.getRange("A2:W2").copyTo(repsheet.getRange(repsheet.getLastRow()+1,1,1,7), {contentsOnly:true});
  sheet1.getRange("A2:W2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('A3:A23').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('B3:B6').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('B11:B23').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A3').setValue('Ctrl+V');
  SpreadsheetApp.getActiveSheet().getRange('D12').setValue('Retrieval');
  SpreadsheetApp.getActiveSheet().getRange('D13').setValue('Yes');
  SpreadsheetApp.getActiveSheet().getRange('D14:D15').setValue('No');
  SpreadsheetApp.getActiveSheet().getRange('D17').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D18').setValue('0');
  SpreadsheetApp.getActiveSheet().getRange('D20').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('G12:G15').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('B7').setValue('(Amount)');
  SpreadsheetApp.getActiveSheet().getRange('B8').setValue('(Trans Date)');
  SpreadsheetApp.getActiveSheet().getRange('B9').setValue('(Case #)');
  SpreadsheetApp.getActiveSheet().getRange('B10').setValue('(Reason)');
  
  }else{
  
  sheet1.getRange("A2:W2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('A3:A23').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('B3:B6').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('B11:B23').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A3').setValue('Ctrl+V');
  SpreadsheetApp.getActiveSheet().getRange('D12').setValue('Retrieval');
  SpreadsheetApp.getActiveSheet().getRange('D13').setValue('Yes');
  SpreadsheetApp.getActiveSheet().getRange('D14:D15').setValue('No');
  SpreadsheetApp.getActiveSheet().getRange('D16').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D17').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('D18').setValue('0');
  SpreadsheetApp.getActiveSheet().getRange('D20').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('B7').setValue('(Amount)');
  SpreadsheetApp.getActiveSheet().getRange('B8').setValue('(Trans Date)');
  SpreadsheetApp.getActiveSheet().getRange('B9').setValue('(Case #)');
  SpreadsheetApp.getActiveSheet().getRange('B10').setValue('(Reason)');

 }
 }
 
function SendToWorkLogAD() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName("AD");
  var sheet2 = ss.getSheetByName("Work Log");
  var repsheet = ss.getSheetByName("Representing");
  var rep = sheet1.getRange("B21");
  var repvalue = rep.getValue();
  
  if (repvalue == 'Yes') {
  
  sheet1.getRange("A2:W2").copyTo(repsheet.getRange(repsheet.getLastRow()+1,1,1,7), {contentsOnly:true});
  sheet1.getRange("A2:W2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
  SpreadsheetApp.getActiveSheet().getRange('A3:F8').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('A3').setValue('Ctrl+SHIFT+V');
  SpreadsheetApp.getActiveSheet().getRange('B18').setValue('Chargeback');
  SpreadsheetApp.getActiveSheet().getRange('B19').setValue('Yes');
  SpreadsheetApp.getActiveSheet().getRange('B20:B21').setValue('No');
  SpreadsheetApp.getActiveSheet().getRange('B22').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('B23').setValue('');
  SpreadsheetApp.getActiveSheet().getRange('B24').setValue('0');
  SpreadsheetApp.getActiveSheet().getRange('B26').setValue('');
  
  }else{
  
  sheet1.getRange("A2:W2").copyTo(sheet2.getRange(sheet2.getLastRow()+1,1,1,7), {contentsOnly:true});
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
