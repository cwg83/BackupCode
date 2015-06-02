function onFormSubmit(e) {
   var date = e.values[0];
   var brand = e.values[1];
   var eGCs = e.values[2];
//Replace google docs line break format with html line breaks
   eGCs = eGCs.replace(/\n/g, '<br>'); 
   
//Send the email   
   var subject = brand + "subject text" + date
   var body    = "Hello " + brand + "," + "<br /><br />" 
   + "Body text here";
   var cc = "";
   MailApp.sendEmail(contactEmail, subject, body, {htmlBody: body, cc: cc});
   }
   
//Custom menu to run a script
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('VLookUp', 'VLookUp')
      .addToUi();
}
function VLookUp() {
  var sh = SpreadsheetApp.getActiveSheet();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var last=ss.getLastRow();
  var data=sh.getRange(1,1,last,2).getValues();// create an array of data from columns A and B
  var valA=Browser.inputBox('Enter value to search in A')
  for(nn=0;nn<data.length;++nn){
    if (data[nn][0]==valA){break} ;// if a match in column B is found, break the loop
      }
Browser.msgBox(data[nn][1]);// show column A
}
