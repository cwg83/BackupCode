function repCredit() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B20').activate();
  spreadsheet.getCurrentCell().setValue('Yes');
  spreadsheet.getRange('B21').activate();
  spreadsheet.getCurrentCell().setValue('Prior Credit');
  spreadsheet.getRange('B22').activate();
  spreadsheet.getCurrentCell().setValue('');
  spreadsheet.getRange('B23').activate();
  spreadsheet.getCurrentCell().setValue('Yes');
  spreadsheet.getRange('B24').activate();  
  spreadsheet.getCurrentCell().setValue('[proof of credit - CyberSource screenshot]');
};

function repUsedP() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B20').activate();
  spreadsheet.getCurrentCell().setValue('Yes');
  spreadsheet.getRange('B21').activate();
  spreadsheet.getCurrentCell().setValue('Product Used');
  spreadsheet.getRange('B22').activate();
  spreadsheet.getCurrentCell().setValue('MP01');
  spreadsheet.getRange('B23').activate();
  spreadsheet.getCurrentCell().setValue('Yes');
  spreadsheet.getRange('B24').activate();  
  spreadsheet.getCurrentCell().setValue('[balance spent notes]');
}; 

function repUsedE() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B20').activate();
  spreadsheet.getCurrentCell().setValue('Yes');
  spreadsheet.getRange('B21').activate();
  spreadsheet.getCurrentCell().setValue('Product Used');
  spreadsheet.getRange('B22').activate();
  spreadsheet.getCurrentCell().setValue('ME01');
  spreadsheet.getRange('B23').activate();
  spreadsheet.getCurrentCell().setValue('Yes');
  spreadsheet.getRange('B24').activate();  
  spreadsheet.getCurrentCell().setValue('[balance spent notes]');
}; 


function repFraudP() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B20').activate();
  spreadsheet.getCurrentCell().setValue('Yes');
  spreadsheet.getRange('B21').activate();
  spreadsheet.getCurrentCell().setValue('Do Not Believe It Is Fraud');
  spreadsheet.getRange('B22').activate();
  spreadsheet.getCurrentCell().setValue('FP01');
  spreadsheet.getRange('B23').activate();
  spreadsheet.getCurrentCell().setValue('Yes');
  spreadsheet.getRange('B24').activate();  
  spreadsheet.getCurrentCell().setValue('');
};

function repFraudE() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B20').activate();
  spreadsheet.getCurrentCell().setValue('Yes');
  spreadsheet.getRange('B21').activate();
  spreadsheet.getCurrentCell().setValue('Do Not Believe It Is Fraud');
  spreadsheet.getRange('B22').activate();
  spreadsheet.getCurrentCell().setValue('FE01');
  spreadsheet.getRange('B23').activate();
  spreadsheet.getCurrentCell().setValue('Yes');
  spreadsheet.getRange('B24').activate();  
  spreadsheet.getCurrentCell().setValue('');
};


function noMismatch() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B20').activate();
  spreadsheet.getCurrentCell().setValue('Yes');
  spreadsheet.getRange('B21').activate();
  spreadsheet.getCurrentCell().setValue('No Financial Or Technical Mismatch');
  spreadsheet.getRange('B22').activate();
  spreadsheet.getCurrentCell().setValue('MP01');
  spreadsheet.getRange('B23').activate();
  spreadsheet.getCurrentCell().setValue('Yes');
  spreadsheet.getRange('B24').activate();  
  spreadsheet.getCurrentCell().setValue('[payment gateway + orders screenshot]');
};
