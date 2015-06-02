function onFormSubmit(e) {
   var brand = e.values[1];
   var eGCs = e.values[2];
//Format today's date as MM/DD/YYYY   
   var today = new Date();
   var dd = today.getDate();
   var mm = today.getMonth()+1; //January is 0!

   var yyyy = today.getFullYear();
    if(dd<10){
        dd='0'+dd
    } 
    if(mm<10){
        mm='0'+mm
    } 
   var today = mm+'/'+dd+'/'+yyyy;
   
   
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var validation = ss.getSheetByName("Contacts");
   var last = ss.getLastRow();
   var data = validation.getRange(1,1,last,2).getValues();// create an array of data from columns A and B
   for(nn=0;nn<data.length;++nn){
    if (data[nn][0]==brand){break} ;// if a match in column B is found, break the loop
      }
   var contactEmail = data[nn][1];
   
   
//Replace google docs line break format with html line breaks
   eGCs = eGCs.replace(/\n/g, '<br>'); 
//Send the email   
   var subject = brand + " eGCs to deactivate " + today
   var body    = "Hello " + brand + "," + "<br /><br />" 
   + "The follow eGift Cards need to be deactivated:"  + "<br /><br />" 
   + eGCs + "<br /><br />" 
   + "Can you please deactivate them and report back the remaining balances?";
   var cc = "";
   MailApp.sendEmail(contactEmail, subject, body, {htmlBody: body, cc: cc});
   }
