function onFormSubmit(e) {
   var brand = e.values[1];
   var eGCs = e.values[2];
   var alias = 'CashStarFailedDeactivations@cashstar.com';
      Logger.log(brand);
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
   
//VLookup - nn=0 is column 1, nn=1 is column 2   
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var validation = ss.getSheetByName("Contacts");
   var last = validation.getLastRow();
   Logger.log(last);
   var data = validation.getRange(1,1,last,2).getValues();// create an array of data from columns A and B
   for(nn=0;nn<data.length;++nn){
    if (data[nn][0]==brand){break} ;// if a match in column B is found, break the loop
      }
   var contactEmail = data[nn][1];
   
//Replace google docs line break format with html line breaks
   eGCs = eGCs.replace(/\n/g, '<br>'); 
   var eGCcount = eGCs.split("<br>").length
         Logger.log(eGCcount);
//Send the email   
   var subject = brand + " eGCs to deactivate " + today
   var cc = "";
   
   var body1 = "Hello " + brand + "," + "<br /><br />" 
   + "The following eGift Card needs deactivation. We were unable to deactivate the eGift Card using our normal means. "
   + "We have refunded the cardholder, but would like to prevent any remaining card balance from being spent. "  + "<br /><br />"
   + eGCs + "<br /><br />" 
   + "Can you please deactivate this eGift Card and report back the remaining balance?";
   
   var body2 = "Hello " + brand + "," + "<br /><br />" 
   + "The following eGift Cards need deactivation. We were unable to deactivate the eGift Cards using our normal means. "
   + "We have refunded the cardholder, but would like to prevent any remaining card balance from being spent. "  + "<br /><br />"
   + eGCs + "<br /><br />" 
   + "Can you please deactivate these eGift Cards and report back the remaining balances?"; 
   
   if (eGCcount == 1) {
   GmailApp.sendEmail(contactEmail, subject, body1, {htmlBody: body1, cc: cc, from: alias})
   return
   }else{
   
   GmailApp.sendEmail(contactEmail, subject, body2, {htmlBody: body2, cc: cc, from: alias})
   return
   }
   }
