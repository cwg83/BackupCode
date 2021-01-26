// Get the worksheets
var user = Session.getActiveUser().getEmail(); // Current user
var ss = SpreadsheetApp.getActiveSpreadsheet(); // Current sheet
var work_log = ss.getSheetByName("Work Log");
var rep_sheet = ss.getSheetByName("Representing");
var updates = ss.getSheetByName("Updates");

// Get some individual cell values
var repvalue = ss.getRange("B20").getValue();
var ordervalue = ss.getRange("B16").getValue();
var brand = ss.getRange("B11").getValue();
var purchase = ss.getRange("B15").getValue();
var stage = ss.getRange("B9").getValue();
var approver = ss.getRange("B18").getValue();
var disputedate = ss.getRange("B4").getValue();

// Establish some ranges
var formulacells = ss.getRange('B3:B15'); // Cells that will need formulas replaced
var textcells = ss.getRange('B16:B27'); // Cells to clear, or to put text back into
var pastesection = ss.getRange('E3:G'); // So we can clear both paste sections for next time
var cbpastecell = ss.getRange('E3'); // So we can reset the paste cell for CB platform data
var orderpastecell = ss.getRange('G3'); // So we can reset the paste cell for order data
var row_for_work_log = ss.getRange("A2:AC2") // Get the data to send to the work log / representing tabs
var extra_pp_cells = ss.getRange("B16:B17")

// Platform formulas
var chase_formulas = [
  ['=IFERROR(INDEX(F35:F84,MATCH("Transaction Date",E35:E84,0))-1,"")'],
  ['=IF(RegExMatch(E6,"Retrieval"),SUBSTITUTE(E7,"Case Status Date",""),IF(E18=" ",SUBSTITUTE(E7,"Case Status Date",""),SUBSTITUTE(E6,"Case Status Date","")))'],
  ['=IFERROR(LEFT(INDEX(F29:F84,MATCH("Amount (Presentment)",E29:E84,0)),LEN(INDEX(F29:F84,MATCH("Amount (Presentment)",E29:E84,0)))-5),"")'],
  ['=IFERROR(LEFT(INDEX(F3:F84,MATCH("Initial Chargeback Amount (Settlement)",E3:E84,0)),LEN(INDEX(F3:F84,MATCH("Initial Chargeback Amount (Settlement)",E3:E84,0)))-5),IFERROR(LEFT(INDEX(F3:F84,MATCH("Initial RR Amount (Presentment)",E3:E84,0)),LEN(INDEX(F3:F84,MATCH("Initial RR Amount (Presentment)",E3:E84,0)))-5),""))'],
  ['=IFERROR(INDEX(F3:F84,MATCH("Method of Payment",E3:E84,0)),"")'],
  ['=IF(A2="","",SUBSTITUTE(INDEX($E$2:$E,MATCH("Chargeback Information",$E$2:$E,0)+1),"Reason Code #: ",""))'],
  ['=IF(A2="","",IF(ISERROR(MATCH("Chargeback",E2:E)),"Retrieval","Chargeback"))'],
  ['=IF(A2="","",VLOOKUP(B8,\'Data Validation\'!V2:W,2,FALSE))'],
  ['=IFERROR(VLOOKUP(INDEX(F35:F84,MATCH("Descriptor",E35:E84,0)),\'Data Validation\'!E2:F,2,FALSE),"")'],
  ['=IFERROR(INDEX(F35:F84,MATCH("Merchant Order #",E35:E84,0)),"")'],
  ['=IF(A2="","",E3)'],
  ['=IF(COUNTIF(F24:F28,"*CAD*")>0,"CAD",IF(COUNTIF(F24:F28,"*GBP*")>0,"GBP","USD"))'],
  ['=IF(B14="USD",B5,SUBSTITUTE(SUBSTITUTE(INDEX(F35:F84,MATCH("Amount (Presentment)",E35:E84,0))," (CAD)",""), "(GBP)",""))']
];

var amex_formulas = [
  ['=IF(E4="","", IF(REGEXMATCH(E8,"CAD"),datevalue((RIGHT(E7,4)&"-"&MID(E7,4,2)&"-"&LEFT(E7,2))),datevalue(E7)))'],
  ['=IF(E4="","",IF(REGEXMATCH(E8,"CAD"),B29&"/"&A29&"/"&C29,E5))'],
  ['=IF(A2="","",IF(REGEXMATCH(E8,"CAD"),RIGHT(E8,LEN(E8)-3),RIGHT(E8,LEN(E8))))'],
  ['=IF(A2="","",IF(REGEXMATCH(E9,"CAD"),RIGHT(E9,LEN(E9)-3),IF(REGEXMATCH(E8,"CAD"),RIGHT(E8,LEN(E8)-3),IF(E9="0.00",E8,E9))))'],
  ['=IF(A2="","","Amex")'],
  ['=IF(E6="Other","Not received",TRIM(E6))'],
  ['=IF($A$2="","",IF($E$13="IQ","Retrieval","Chargeback"))'],
  ['=IF(A2="","",VLOOKUP(B8,\'Data Validation\'!V2:W,2,FALSE))'],
  ['=IF(A2="","",VLOOKUP(E10,\'Data Validation\'!D2:F,3,FALSE))'],
  ['=E11'],
  ['=IF(A2="","",E3)'],
  ['=IF(REGEXMATCH(E8,"CAD"),"CAD","USD")'],
  ['=IF(A2="","",SUBSTITUTE(E8,"CAD",""))']
];

var pp_formulas = [
  ['=DATEVALUE(INDEX($G$3:$G,MATCH("Date:",$G$3:$G,0)+2))'],
  ['=IFERROR(INDEX(E2:E30,MATCH("Date reported",E2:E30,0)+1),IFERROR(DATEVALUE(VLOOKUP("Date of Complaint:",E3:F15,2,FALSE)),C27))'],
  ['=IFERROR(IFERROR(REGEXREPLACE(VLOOKUP("Transaction Amount:",E2:F12,2,FALSE),"[^0-9.]",""),REGEXREPLACE(INDEX(E2:E30,MATCH("Transaction amount",E2:E30,0)+1),"[^0-9.]","")),SUBSTITUTE(INDEX(E2:E30,MATCH("Transaction amount*",E2:E30,0)),"Transaction amount$",""))'],
  ['=B5'],
  ['="PayPal"'],
  ['=IFERROR(IFERROR(SUBSTITUTE(INDEX(E2:E,MATCH("Dispute reason*",E2:E,0)),"Dispute reason",""),VLOOKUP("Reason for Dispute:",E2:F12,2,FALSE)),"")'],
  ['=IF($B$8="Unauthorized Payment","Chargeback",IF($B$8="Merchandise","Chargeback",IF($B$8="Item not received","Chargeback","Retrieval")))'],
  ['=IFERROR(VLOOKUP($B$8,\'Data Validation\'!$V2:W,2,FALSE),"")'],
  ['=iferror(REGEXEXTRACT($J$3,"–~U~(.+?)~U~"),"")'],
  ['=IFERROR(IF(REGEXMATCH(E3,"Transaction"),VLOOKUP("Invoice ID:",E2:F12,2,FALSE),SUBSTITUTE(INDEX(E2:E,MATCH("Invoice ID*",E2:E,0)),"Invoice ID","")),(INDEX(E2:E,MATCH("Invoice ID",E2:E,0)+1)))'],
  ['=IFERROR(IF(REGEXMATCH($E$3,"PP-D"),SUBSTITUTE($E$3,"Case ID: ",""),VLOOKUP("PayPal Case ID:",E3:F15,2,FALSE)),"")'],  
  ['="USD"'],
  ['=$B$5'],          
];

var extra_pp_formulas = [

  ['=IFERROR(SUBSTITUTE(INDEX($G$3:$G,MATCH("Risk score*",$G$3:$G,0)-1),"ORDER ",""),"")'],
  ['=SPLIT(LOWER(INDEX($G3:G,MATCH("Current Balance:",$G3:G,0)+2)),"abcdefhijklmnopqrstuvwxyz")']  
];

var ad_formulas = [
  ['=IF(E4="","",LEFT(I4,LEN(I4)-13))'],
  ['=IF(E4="","",DATEVALUE(LEFT(I3,LEN(I3)-13)))'],
  ['=IF(A2="","",RIGHT(I6,LEN(I6)-4))'],
  ['=IF(A2="","",RIGHT(I5,LEN(I5)-4))'],
  ['=IF(A2="","",F6)'],
  ['=TRIM(F3)'],
  ['=IF(A2="","","Chargeback")'],
  ['=IF(A2="","",VLOOKUP(B8,\'Data Validation\'!U2:V,2,FALSE))'],
  ['=IFERROR(VLOOKUP("Payment Merchant Reference",E3:I,2,False),VLOOKUP("Payment Merchant Reference",H3:I,2,False))'],         
  ['=VLOOKUP("Payment Psp Reference",E3:F,2,False)'],
  ['=IF(A2="","",IF(RegExMatch(I6,"GBP"),"GBP",IF(RegExMatch(I6,"HKD"),"HKD",IF(RegExMatch(I6,"EUR"),"EUR"))))'],
  [''],
  ['=IF($A$2="","",$B$5)'],       
];


// Text range.
var text = [
  [''],
  ['0'],
  ['Score Autoreleased'],
  ['No'],
  ['No'],
  [''],
  [''],
  [''],
  [''],
  [''],
  [''],
  ['']
];

//Set the platform manually, or when resetting the sheet    
function setPlatform()   {

  platform = ss.getRange("A30").getValue();      
  // Set the formulas array depending on which platform is selected

  switch(platform)  {
    case "Chase":
      formulas = chase_formulas;
      break;
    case "Amex":
      formulas = amex_formulas;
      break;
    case "PayPal":
      formulas = pp_formulas;
      extra_pp_cells.setFormulas(extra_pp_formulas);
      break;
    case "Adyen":
      formulas = ad_formulas;
      break;
  }
  formulacells.setFormulas(formulas);

}

// Send it to the WERK LOG and to THA REP REZENT tab, YO
function sendToWorkLog() {


  ss.getRange('B28').setValue(user); // set the current user

  // Check for missing entries

  if (ordervalue.indexOf('CN') === -1) {
    Browser.msgBox("Please enter a valid Order #");
    return
  } else if (brand == "") {
    Browser.msgBox("Please make sure the 'Brand' field is populated");
    return
  } else if (brand == "#N/A") {
    Browser.msgBox("This brand needs to be added to the Data Validation sheet");
    return
  } else if (purchase == "") {
    Browser.msgBox("Please make sure the 'Purchase Amount' field is populated");
    return
  } else if (approver == "") {
    Browser.msgBox("Please make sure the 'Approver' field is populated");
    return
  } else if (disputedate == "") {
    Browser.msgBox("Please make sure the 'Dispute Date' field is populated");
    return
  }

  // Add the row of collected data to the Representing tab, if representing
  if (repvalue == 'Yes') {
    row_for_work_log.copyTo(rep_sheet.getRange(rep_sheet.getLastRow() + 1, 1, 1, 7), {
      contentsOnly: true
    });
    var replastRow = rep_sheet.getLastRow(); // Find last row of the Representing tab
    rep_sheet.insertRowAfter(replastRow); // Append a blank row to the end of the the Representing tab
  }

  // Add the row to the work log
  row_for_work_log.copyTo(work_log.getRange(work_log.getLastRow() + 1, 1, 1, 7), {
    contentsOnly: true
  });
  var lastRow = work_log.getLastRow();
  work_log.insertRowAfter(lastRow); //Append a blank row to the end of the WorkLog


  // Reset the formulas and text cells

  extra_pp_cells.setValues([[""],[""]]);
  setPlatform();
  formulacells.setFormulas(formulas);
  textcells.setValues(text);
  pastesection.setValue('');
  cbpastecell.setValue('CB Paste'); //Paste Cell
  orderpastecell.setValue('Order Paste'); //Paste Cell      

  var range = work_log.getRange('G2:K');

  // Money format
  range.setNumberFormat("#,##0.00;$(#,##0.00)");
}








// Define representment document generation formulas
var docgen_formulas = [
['=iferror(REGEXEXTRACT($J$3,"Name:~U~(.+?)~U~"),"")'],
['=iferror(REGEXEXTRACT($J$3,"Address:~U~(.+?)~U~"),"")'],
['=iferror(REGEXEXTRACT($J$3,I4&"~U~(.+?)~U~"),"")'],
['=iferror(REGEXEXTRACT($J$3,I5&"~U~(.+?)~U~"),"")'],
['=iferror(REGEXEXTRACT($J$3,I6&"~U~(.+?)~U~"),"")'],
['=IF(A30="PayPal","N/A",iferror(REGEXEXTRACT($J$3,"CC#:~U~(.+?)~U~"),"N/A"))'],
['=IF(I8="N/A","N/A",iferror(REGEXEXTRACT($J$3,"AVS Response Detail:~U~(.+?)~U~"),"N/A"))'],
['=IF($G$3="","",IF($I$8="N/A","N/A",INDEX($G$3:$G,MATCH("CC Auth Details:",$G$3:$G,0)+2)))'],
['=IF(I8="N/A","N/A",iferror(REGEXEXTRACT($J$3,"Card Code Response:~U~(.+?)~U~"),"N/A"))'],
['=DATEVALUE(REGEXEXTRACT($J$3,"Date:~U~(.+?)~U~"))'],
['=iferror(if(REGEXMATCH($J$3, "Recipient Name:~U~Multiple")=TRUE, REGEXEXTRACT(REGEXEXTRACT($J$3, "Recipient Name:~U~Multiple(.+)"), "Recipient Name:~U~(.+?)~U~"),REGEXEXTRACT($J$3, "Recipient Name:~U~(.+?)~U~")),"")'],
['=iferror(if(REGEXMATCH($J$3, "Recipient Email:~U~Multiple")=TRUE, REGEXEXTRACT(REGEXEXTRACT($J$3, "Recipient Email:~U~Multiple(.+)"), "Recipient Email:~U~(.+?)~U~"),REGEXEXTRACT($J$3, "Recipient Email:~U~(.+?)~U~")),"")'],
['=REGEXEXTRACT($J$3,"Purchaser IP Address:~U~(.+?)~U~")'],
['=REGEXEXTRACT($J$3,"Device ID:~U~(.+?)~U~")'],
['=REGEXEXTRACT($J$3,"Web Request ID:~U~(.+?)~U~")'],
['=REGEXEXTRACT($J$3,"Purchase Total:~U~(.+?)~U~")'],
['=REGEXEXTRACT($J$3,"Score:~U~(.+?)~U~")'],
['=IF(A30="PayPal","PayPal",REGEXEXTRACT($J$3,"Bank:~U~(.+?)~U~"))'],
['=REGEXEXTRACT($J$3,"Region:~U~(.+?)~U~")'],
['=REGEXEXTRACT($J$3,"City:~U~(.+?)~U~")'],
['=REGEXEXTRACT($J$3,"Country Code:~U~(.+?)~U~")'],
['=REGEXEXTRACT($J$3,"ORG:~U~(.+?)~U~")'],
['=REGEXEXTRACT($J$3,"Distance:~U~(.+?)~U~")'],
['=REGEXEXTRACT($J$3,"Proxy Score:~U~(.+?)~U~")'],
['=REGEXEXTRACT($J$3,"Type:~U~(.+?)~U~")'],
['=B17'],
['=IF("Redemption URL:"=REGEXEXTRACT($J$3,"Viewed IP:~U~(.+?)~U~"),"N/A","")'],
['=IF(I29="N/A",I29,REGEXEXTRACT($J$3,"Viewed Date:~U~(.+?)~U~"))'],
['=iferror(REGEXEXTRACT($K$4,"Address:~U~(.+?)~U~"),"")'],
['=iferror(REGEXEXTRACT($K$4,"Address:~U~(.+?)~U~"),"")'],
['=iferror(REGEXEXTRACT($K$4,I32&"~U~(.+?)~U~"),"")'],
['=iferror(REGEXEXTRACT($K$4,I33&"~U~(.+?)~U~"),"")'],
['=IF("Address:"=REGEXEXTRACT($J$3,"Tracking #:~U~(.+?)~U~"),"N/A",iferror(REGEXEXTRACT($J$3,"Tracking #:~U~(.+?)~U~"),""))'],
['=iferror(REGEXEXTRACT($J$3,"Message:~U~(.+?)~U~"),"")'],
['=iferror(REGEXEXTRACT($J$3,"–~U~(.+?)~U~"),"")'],
['=B13'],
['=iferror(index(split(REGEXEXTRACT(REGEXEXTRACT($J$3, "CUSTOMER: (.+?)Payment History"), "~U~(.+?)~U~")," "),0,1),"")'],
['=B21'],
['=B22'],
['=IF(G3="Order Paste","",C8)'],
]; 


// Get the order data
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var purchaser_name = ss.getRange("I3").getValue();
   var billing1 = ss.getRange("I4").getValue();
   var billing2 = ss.getRange("I5").getValue();
   var billing3 = ss.getRange("I6").getValue();
   var purchaser_phone = ss.getRange("I7").getValue();
   var last4 = ss.getRange("I8").getValue();
   var avs_response = ss.getRange("I9").getValue();
   var cc_auth = ss.getRange("I10").getValue();
   var cc_response = ss.getRange("I11").getValue();
   var purchase_date = ss.getRange("I12").getValue();
   var recip_name = ss.getRange("I13").getValue();
   var recip_email = ss.getRange("I14").getValue();
   var purchaser_IP = ss.getRange("I15").getValue();
   var device_ID = ss.getRange("I16").getValue();
   var webrequest_ID = ss.getRange("I17").getValue();
   var purchase_amount = ss.getRange("I18").getValue();
   var minfraud_score = ss.getRange("I19").getValue();
   var bank_name = ss.getRange("I20").getValue();
   var ip_region = ss.getRange("I21").getValue();
   var ip_city = ss.getRange("I22").getValue();
   var ip_country = ss.getRange("I23").getValue();
   var purchaser_ISP = ss.getRange("I24").getValue();
   var ip_distance = ss.getRange("I25").getValue();
   var proxy_score = ss.getRange("I26").getValue();
   var giftcard_type = ss.getRange("I27").getValue(); 
   var balance = ss.getRange("I28").getValue();   
   var viewed_IP = ss.getRange("I29").getValue();
   var viewed_date = ss.getRange("I30").getValue();
   var shipping1 = ss.getRange("I31").getValue();  
   var shipping2 = ss.getRange("I32").getValue();  
   var shipping3 = ss.getRange("I33").getValue();   
   var shipping4 = ss.getRange("I34").getValue();   
   var tracking = ss.getRange("I35").getValue();   
   var message = ss.getRange("I36").getValue();      
   var brand1 = ss.getRange("I37").getValue();     
   var case_number = ss.getRange("I38").getValue();   
   var purchaser_email = ss.getRange("I39").getValue();
   var rep_reason = ss.getRange("I40").getValue();

   var alias = 'paymentsteam@cashstar.com';



  function createDoc() {

  // Template Doc IDs ******************************************


  var ME01docTemplate = "1aIGrZfMJLCIkD10jlZxvOJEs6uEn9fA7Y9Xd8zO7DgE";  // *** Merchandise / E-Gift Card / V1 ***
  var FE01docTemplate = "1zbehfU7XqRNG5qtsJdGJ-5RQgYK7iTGH3nUr4WkpLAo";  // *** Fraud / E-Gift Card / V1 ***
  var MP01docTemplate = "1KbIEymdoOOwvfp1Rm423j9FrnHMUX21kZQQasg46xJo";  // *** Merchandise / Plastic Gift Card / V1 ***
  var FP01docTemplate = "1kmFBdtt2k0Jc9xXjScpYJZzyC9b2jEjRL6cJTbwNXdA";  // *** Fraud / Plastic Gift Card / V1 ***
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wch_doc = ss.getRange("I41").getValue();
  var notes = ss.getRange("I42").getValue();

  combinedNotes = notes;

  // Determine which doc is to be generated and pass the specific notes to it  

  switch(wch_doc) {
    case "ME01":
      makeDoc(ME01docTemplate);
      break;
    case "FE01":
      makeDoc(FE01docTemplate);     
      break;
    case "MP01":
      makeDoc(MP01docTemplate);
      break;
    case "FP01":
      makeDoc(FP01docTemplate);
      break;
  }

// Generate the cover letter
  makeCover();

// Clear input cells and reset fomulas
 var docgen_formulacells = ss.getRange('I3:I42');  
 var docgen_pastesection = ss.getRange('G3:G');

 docgen_pastesection.setValue('');

 docgen_formulacells.setFormulas(docgen_formulas);  

// ******************************************

}  

// Placeholder replacement function

function fillInText(copyBody) {

// Replace place holder keys
   copyBody.replaceText('keyNotes', combinedNotes);   
   copyBody.replaceText('keyPurchaserName', purchaser_name);
   copyBody.replaceText('keyPurchaserName', purchaser_name);
   copyBody.replaceText('keyBilling1', billing1);
   copyBody.replaceText('keyBilling2', billing2);
   copyBody.replaceText('keyBilling3', billing3);
   copyBody.replaceText('keyPhone', purchaser_phone);
   copyBody.replaceText('keyLast4', last4);
   copyBody.replaceText('keyAVS', avs_response);
   copyBody.replaceText('keyCCAuth', cc_auth);
   copyBody.replaceText('keyCCResponse', cc_response);
   copyBody.replaceText('keyPurchaseDate', purchase_date);
   copyBody.replaceText('keyRecipName', recip_name);
   copyBody.replaceText('keyRecipEmail', recip_email);
   copyBody.replaceText('keyPurchaserIP', purchaser_IP);
   copyBody.replaceText('keyDeviceID', device_ID);
   copyBody.replaceText('keyWebRequest', webrequest_ID);
   copyBody.replaceText('keyPurchaseAmount', purchase_amount);
   copyBody.replaceText('keyMinFraud', minfraud_score);
   copyBody.replaceText('keyBank', bank_name);
   copyBody.replaceText('keyIPRegion', ip_region);
   copyBody.replaceText('keyIPCity', ip_city);
   copyBody.replaceText('keyPurchaserIP', purchaser_ISP);
   copyBody.replaceText('keyIPDistance', ip_distance);
   copyBody.replaceText('keyProxyScore', proxy_score);
   copyBody.replaceText('keyBrand', brand1);
   copyBody.replaceText('keyBalance', balance);
   copyBody.replaceText('keyPurchaserEmail', purchaser_email);
   copyBody.replaceText('keyViewIP', viewed_IP);
   copyBody.replaceText('keyViewDate', viewed_date);
   copyBody.replaceText('keyRepreason', rep_reason);   
   copyBody.replaceText('keyCase#', case_number);
   copyBody.replaceText('keyShipping1', shipping1);
   copyBody.replaceText('keyShipping2', shipping2);
   copyBody.replaceText('keyShipping3', shipping3);
   copyBody.replaceText('keyShipping4', shipping4);
   copyBody.replaceText('keyTracking', tracking);
   copyBody.replaceText('keyType', giftcard_type);
   copyBody.replaceText('keyMessage', message);   

   var todaysDate = Utilities.formatDate(new Date(), "GMT", "MM/dd/yyyy"); 
   copyBody.replaceText('keyTodaysDate', todaysDate);

}


  // Function to generate the cover letter

  function makeCover() {   
     // Get document template, copy it as a new temp doc, and save the Doc’s id
     var copyId = DriveApp.getFileById("16XBDN0NFbMVlw0cKLBwK8DUeRBQ7G5s58jxAVaPXpqI")
     .makeCopy(case_number+"_cover")
     .getId();

     // Open the temporary document
     var copyDoc = DocumentApp.openById(copyId);

     // Get the document’s body section and fill in values from the sheet
     var copyBody = copyDoc.getActiveSection();
     fillInText(copyBody);

     // Save and close the temporary document
     copyDoc.saveAndClose();

     // Convert temporary document to PDF by using the getAs blob conversion
     var pdf = DriveApp.getFileById(copyId).getAs("application/pdf");

     // Save the PDF and trash the temp file
     var fid = '0B3SJLpqm5CXcR1g5U3l6dmhhMGc';
     var folder = DriveApp.getFolderById(fid);
     folder.createFile(pdf); 
     DriveApp.getFileById(copyId).setTrashed(true);


}

// Function to generate the different doc types 

  function makeDoc(docTemplate) {
    // Get document template, copy it as a new temp doc, and save the Doc’s id
    var copyId = DriveApp.getFileById(docTemplate)
    .makeCopy(case_number)
    .getId();

    // Open the temporary document
    var copyDoc = DocumentApp.openById(copyId);

    // Get the document’s body section and fill in values from the sheet
    var copyBody = copyDoc.getActiveSection();
    fillInText(copyBody);

    // Save and close the temporary document
    copyDoc.saveAndClose();

    // Convert temporary document to PDF by using the getAs blob conversion
    var pdf = DriveApp.getFileById(copyId).getAs("application/pdf");

    // Save the PDF and trash the temp file
    var fid = '0B3SJLpqm5CXcR1g5U3l6dmhhMGc';
    var folder = DriveApp.getFolderById(fid);
    folder.createFile(pdf); 
    DriveApp.getFileById(copyId).setTrashed(true);
  }
