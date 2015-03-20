// Global variables 
var docTemplate = "15kvC3M8b0Me3Vi3GwErQvG60BlTC0qAHMXrlc3Ocky8";  // *** replace with your template ID ***
var docName     = "RefundByCheck"

function onFormSubmit(e) { // add an onsubmit trigger
// Full name and email address values come from the spreadsheet form
   var first_name = e.values[1];
   var last_name = e.values[2];
   var customer_email = e.values[3];
   var brand = e.values[4];
   var amount = e.values[5];
   var purchase_date = e.values[6];
   var customer_address = e.values[7];
   var rep_name = e.values[8];
   var order_number = e.values[9];
// Get document template, copy it as a new temp doc, and save the Doc’s id
   var copyId = DocsList.getFileById(docTemplate)
                .makeCopy(docName+'_'+order_number)
                .getId();
// Open the temporary document
   var copyDoc = DocumentApp.openById(copyId);
// Get the document’s body section
   var copyBody = copyDoc.getActiveSection();
// Replace place holder keys,  
   copyBody.replaceText('keyFirst', first_name);
   copyBody.replaceText('keyLast', last_name);
   copyBody.replaceText('keyBrand', brand);
   copyBody.replaceText('keyAmount', amount);
   copyBody.replaceText('keyPurchaseDate', purchase_date);
   copyBody.replaceText('keyAddress', customer_address);
   copyBody.replaceText('keyRep', rep_name);
   copyBody.replaceText('keyOrder', order_number);
   var todaysDate = Utilities.formatDate(new Date(), "GMT", "MM/dd/yyyy"); 
   copyBody.replaceText('keyTodaysDate', todaysDate);
// Save and close the temporary document
   copyDoc.saveAndClose();
// Convert temporary document to PDF by using the getAs blob conversion
   var pdf = DocsList.getFileById(copyId).getAs("application/pdf");
   var folder = DocsList.getFolder('Refunds by Check');
   var movefile = DocsList.createFile(pdf);
   movefile.addToFolder(folder);
   movefile.removeFromFolder(DocsList.getRootFolder());
// Attach PDF and send the email
   var subject = "Refund by Check regarding Order Number " + order_number
   var body    = "Hello " + first_name + " " + last_name + "," + "<br /><br />" 
   + "Thank you for contacting our support team." + "<br /><br />"
   + "Due to the amount of time that has elapsed since the purchase date, we are unable to refund you directly to your original form of payment. "
   + "However, we can issue a refund check by U.S. mail. I've included our standard consent form for this process." + "<br /><br />"
   + "Please verify the address provided in the first paragraph as that is where the Postal Service will attempt to deliver the check. "
   + "Corrections to the address may be made in the area below the signature area if necessary. If everything is correct, please sign and return the form to me by email, fax, or mail. "
   + "I will then have my accounting group issue you a check right away." + "<br /><br />"
   + "Regards, " + "<br /><br />"
   + rep_name + ", Payments Department" + "<br />"
   + "paymentsteam@cashstar.com";
   var cc = "paymentsteam@cashstar.com";
   MailApp.sendEmail(customer_email, subject, body, {htmlBody: body, attachments: pdf, cc: cc}); 
   DocsList.getFileById(copyId).setTrashed(true);
// Delete temp file
}
