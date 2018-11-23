function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Macros')
      .addItem('Prior Credit', 'repCredit')
      .addItem('No Financial Mismatch', 'noMismatch')
  
      .addSeparator()
      
      .addSubMenu(ui.createMenu('Plastic')
          .addItem('Do Not Believe It Is Fraud', 'repFraudP')            
          .addItem('Product Used', 'repUsedP'))          
          
      .addSubMenu(ui.createMenu('Electronic')
          .addItem('Do Not Believe It Is Fraud', 'repFraudE')             
          .addItem('Product Used', 'repUsedE'))
          
      .addToUi();
}
	
