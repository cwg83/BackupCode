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
      var approver = sheet1.getRange("B18");


      SpreadsheetApp.getActiveSheet().getRange('B27').setValue(user);

      if (ordervalue.indexOf('CNZ') === -1) {
          Browser.msgBox("Please enter a valid Order #");
          return
      } else if (brand.getValue() == "") {
          Browser.msgBox("Please make sure the 'Brand' field is populated");
          return
      } else if (brand.getValue() == "#N/A") {
          Browser.msgBox("This brand needs to be added to the Data Validation sheet");
          return
      } else if (purchase.getValue() == "") {
          Browser.msgBox("Please make sure the 'Purchase Amount' field is populated");
          return
      } else if (approver.getValue() == "") {
          Browser.msgBox("Please make sure the 'Approver' field is populated");
          return
      }
      if (stage.getValue() == 'Graduated Chargeback' || stage.getValue() == 'Second Chargeback' || stage.getValue() == 'Pre-Arbitration') {
          sheet1.getRange("A2:Y2").copyTo(updates.getRange(updates.getLastRow() + 1, 1, 1, 7), {
              contentsOnly: true
          });
      }
      if (repvalue == 'Yes') {
          sheet1.getRange("A2:Y2").copyTo(repsheet.getRange(repsheet.getLastRow() + 1, 1, 1, 7), {
              contentsOnly: true
          });
          var replastRow = repsheet.getLastRow(); //Find last row of RepSheet
          repsheet.insertRowAfter(replastRow); //Append a blank row to the end of the repsheet
      }
      if (stage.getValue() == 'Retrieval' || stage.getValue() == 'Chargeback') {
          sheet1.getRange("A2:Y2").copyTo(sheet2.getRange(sheet2.getLastRow() + 1, 1, 1, 7), {
              contentsOnly: true
          });
      }
      SpreadsheetApp.getActiveSheet().getRange('B11').setFormula("=IF(A2=\"\",\"\",VLOOKUP($D$18,'Data Validation'!$C$2:$F,4,FALSE))"); //Brand
      SpreadsheetApp.getActiveSheet().getRange('B9').setFormula("=IF($A$2=\"\",\"\",IF(RegExMatch($E$25,\"First Chargeback\"),\"Chargeback\",IF(RegExMatch($E$25,\"Second Chargeback\"),\"Second Chargeback\",IF(RegExMatch($E$25,\"Pre-Arbitration\"),\"Pre-Arbitration\",IF(RegExMatch($E$26,\"Retrieval\"),\"Retrieval\",\"\")))))"); //Stage Reached
      SpreadsheetApp.getActiveSheet().getRange('B16').setValue(''); //Order #
      SpreadsheetApp.getActiveSheet().getRange('B14').setFormula("=IF(A2=\"\",\"\",IF(REGEXMATCH($E$10,\"CAD\"),\"CAD\",\"USD\"))"); //Currency
      SpreadsheetApp.getActiveSheet().getRange('B15').setFormula("=IF($B$14=\"USD\",TRIM($B$5),\"\")"); //Purchase Amount
      SpreadsheetApp.getActiveSheet().getRange('B18').setValue('Score Autoreleased'); //Approver
      SpreadsheetApp.getActiveSheet().getRange('B19:B20').setValue('No'); //Verified | Represented
      SpreadsheetApp.getActiveSheet().getRange('B21:B26').setValue(''); //Rep Reason --> Documentation 4
      SpreadsheetApp.getActiveSheet().getRange('B17').setValue('0'); //Balance
      SpreadsheetApp.getActiveSheet().getRange('E3:F39').setValue(''); //Paste Section
      SpreadsheetApp.getActiveSheet().getRange('E3').setValue('Ctrl+SHIFT+V'); //Paste Cell

      var lastRow = sheet2.getLastRow(); //Find last row of WorkLog
      sheet2.insertRowAfter(lastRow); //Append a blank row to the end of the WorkLog
  }

  function SendToWorkLogCH() { //This is the script for the Chase sheet

      var sheet1 = ss.getSheetByName("CH");
      var rep = sheet1.getRange("B20");
      var repvalue = rep.getValue();
      var order = sheet1.getRange("B16");
      var ordervalue = order.getValue();
      var brand = sheet1.getRange("B11");
      var purchase = sheet1.getRange("B15");
      var stage = sheet1.getRange("B9");
      var approver = sheet1.getRange("B18");
      var disputedate = sheet1.getRange("B4");

      SpreadsheetApp.getActiveSheet().getRange('B27').setValue(user);

      if (ordervalue.indexOf('CNZ') === -1) {
          Browser.msgBox("Please enter a valid Order #");
          return
      } else if (brand.getValue() == "") {
          Browser.msgBox("Please make sure the 'Brand' field is populated");
          return
      } else if (brand.getValue() == "#N/A") {
          Browser.msgBox("This brand needs to be added to the Data Validation sheet");
          return
      } else if (purchase.getValue() == "") {
          Browser.msgBox("Please make sure the 'Purchase Amount' field is populated");
          return
      } else if (approver.getValue() == "") {
          Browser.msgBox("Please make sure the 'Approver' field is populated");
          return
      } else if (disputedate.getValue() == "") {
          Browser.msgBox("Please make sure the 'Dispute Date' field is populated");
          return
      }
      if (stage.getValue() == 'Graduated Chargeback' || stage.getValue() == 'Second Chargeback' || stage.getValue() == 'Pre-Arbitration') {
          sheet1.getRange("A2:Y2").copyTo(updates.getRange(updates.getLastRow() + 1, 1, 1, 7), {
              contentsOnly: true
          });
      }
      if (repvalue == 'Yes') {
          sheet1.getRange("A2:Y2").copyTo(repsheet.getRange(repsheet.getLastRow() + 1, 1, 1, 7), {
              contentsOnly: true
          });
          var replastRow = repsheet.getLastRow(); //Find last row of RepSheet
          repsheet.insertRowAfter(replastRow); //Append a blank row to the end of the repsheet
      }
      if (stage.getValue() == 'Retrieval' || stage.getValue() == 'Chargeback') {
          sheet1.getRange("A2:Y2").copyTo(sheet2.getRange(sheet2.getLastRow() + 1, 1, 1, 7), {
              contentsOnly: true
          });
      }
      SpreadsheetApp.getActiveSheet().getRange('B9').setFormula("=IF(A2=\"\",\"\",IF(E17=\"\",IF(RegExMatch($E$18,\"Chargeback\"),\"Chargeback\",\"Retrieval\"),IF(RegExMatch($E$19,\"Chargeback\"),\"Chargeback\",\"Retrieval\")))"); //Stage Reached
      SpreadsheetApp.getActiveSheet().getRange('B11').setFormula("=IF(A2=\"\",\"\",IF(E17=\"\",VLOOKUP($F$71,'Data Validation'!$E$2:$F,2,FALSE),VLOOKUP($F$72,'Data Validation'!$E$2:$F,2,FALSE)))"); //Brand  
      SpreadsheetApp.getActiveSheet().getRange('B14').setFormula("=IF(REGEXMATCH($F$26,\"CAD\"),\"CAD\",IF(REGEXMATCH($F$25,\"CAD\"),\"CAD\",IF(REGEXMATCH($F$27,\"CAD\"),\"CAD\",\"USD\")))"); //Currency
      SpreadsheetApp.getActiveSheet().getRange('B15').setFormula("=IF(B14=\"USD\",B5,SUBSTITUTE($F$27,\" (CAD)\",\"\"))"); //Purchase Amount  
      SpreadsheetApp.getActiveSheet().getRange('B16').setValue(''); //Order #
      SpreadsheetApp.getActiveSheet().getRange('B17').setValue('0'); //Balance 
      SpreadsheetApp.getActiveSheet().getRange('B18').setValue('Score Autoreleased'); //Approver
      SpreadsheetApp.getActiveSheet().getRange('B19:B20').setValue('No'); //Verified | Represented
      SpreadsheetApp.getActiveSheet().getRange('B21:B26').setValue(''); //Rep Reason --> Documentation 4
      SpreadsheetApp.getActiveSheet().getRange('E3:F83').setValue(''); //Paste Section
      SpreadsheetApp.getActiveSheet().getRange('E3').setValue('Ctrl+SHIFT+V'); //Paste Cell

      var lastRow = sheet2.getLastRow(); //Find last row of WorkLog
      sheet2.insertRowAfter(lastRow); //Append a blank row to the end of the WorkLog
  }

  function SendToWorkLogAM() { //This is the script for the Amex sheet

      var sheet1 = ss.getSheetByName("AM");
      var rep = sheet1.getRange("B20");
      var repvalue = rep.getValue();
      var order = sheet1.getRange("B16");
      var ordervalue = order.getValue();
      var brand = sheet1.getRange("B11");
      var purchase = sheet1.getRange("B15");
      var approver = sheet1.getRange("B18");

      SpreadsheetApp.getActiveSheet().getRange('B27').setValue(user);

      if (ordervalue.indexOf('CNZ') === -1) {
          Browser.msgBox("Please enter a valid Order #");
          return
      } else if (brand.getValue() == "") {
          Browser.msgBox("Please make sure the 'Brand' field is populated");
          return
      } else if (brand.getValue() == "#N/A") {
          Browser.msgBox("This brand needs to be added to the Data Validation sheet");
          return
      } else if (purchase.getValue() == "") {
          Browser.msgBox("Please make sure the 'Purchase Amount' field is populated");
          return
      } else if (approver.getValue() == "") {
          Browser.msgBox("Please make sure the 'Approver' field is populated");
          return
      }

      if (repvalue == 'Yes') {

          sheet1.getRange("A2:Y2").copyTo(repsheet.getRange(repsheet.getLastRow() + 1, 1, 1, 7), {
              contentsOnly: true
          });
          var replastRow = repsheet.getLastRow(); //Find last row of RepSheet
          repsheet.insertRowAfter(replastRow); //Append a blank row to the end of the repsheet
      }

      sheet1.getRange("A2:Y2").copyTo(sheet2.getRange(sheet2.getLastRow() + 1, 1, 1, 7), {
          contentsOnly: true
      });
      SpreadsheetApp.getActiveSheet().getRange('B11').setFormula("=IF(A2=\"\",\"\",VLOOKUP(D18,'Data Validation'!D2:F,3,FALSE))"); //Brand
      SpreadsheetApp.getActiveSheet().getRange('B9').setFormula("=IF($A$2=\"\",\"\",IF($E$13=\"IQ\",\"Retrieval\",\"Chargeback\"))"); //Stage Reached
      SpreadsheetApp.getActiveSheet().getRange('B16').setValue(''); //Order #
      SpreadsheetApp.getActiveSheet().getRange('B14').setValue('USD'); //Currency
      SpreadsheetApp.getActiveSheet().getRange('B18').setValue('Score Autoreleased'); //Approver
      SpreadsheetApp.getActiveSheet().getRange('B19:B20').setValue('No'); //Verification | Represented
      SpreadsheetApp.getActiveSheet().getRange('B21:B26').setValue(''); //Rep Reason --> Documentation 4
      SpreadsheetApp.getActiveSheet().getRange('B17').setValue('0'); //Balance
      SpreadsheetApp.getActiveSheet().getRange('E3:F13').setValue(''); //Paste area
      SpreadsheetApp.getActiveSheet().getRange('E3').setValue('Ctrl+SHIFT+V'); //Paste cell

      var lastRow = sheet2.getLastRow(); //Find last row of WorkLog
      sheet2.insertRowAfter(lastRow); //Append a blank row to the end of the WorkLog
  }

  function SendToWorkLogPP() { //This is the script for the PayPal sheet

      var sheet1 = ss.getSheetByName("PP");
      var rep = sheet1.getRange("B20");
      var repvalue = rep.getValue();
      var order = sheet1.getRange("B16");
      var ordervalue = order.getValue();
      var brand = sheet1.getRange("B15");
      var purchase = sheet1.getRange("B14");
      var approver = sheet1.getRange("B18");

      SpreadsheetApp.getActiveSheet().getRange('B27').setValue(user);

      if (ordervalue.indexOf('CNZ') === -1) {
          Browser.msgBox("Please enter a valid Order #");
          return
      } else if (brand.getValue() == "") {
          Browser.msgBox("Please make sure the 'Brand' field is populated");
          return
      } else if (brand.getValue() == "#N/A") {
          Browser.msgBox("This brand needs to be added to the Data Validation sheet");
          return
      } else if (purchase.getValue() == "") {
          Browser.msgBox("Please make sure the 'Purchase Amount' field is populated");
          return
      } else if (approver.getValue() == "") {
          Browser.msgBox("Please make sure the 'Approver' field is populated");
          return
      }

      if (repvalue == 'Yes') {

          sheet1.getRange("A2:Y2").copyTo(repsheet.getRange(repsheet.getLastRow() + 1, 1, 1, 7), {
              contentsOnly: true
          });
          var replastRow = repsheet.getLastRow(); //Find last row of RepSheet
          repsheet.insertRowAfter(replastRow); //Append a blank row to the end of the repsheet 
      }

      sheet1.getRange("A2:Y2").copyTo(sheet2.getRange(sheet2.getLastRow() + 1, 1, 1, 7), {
          contentsOnly: true
      });
      SpreadsheetApp.getActiveSheet().getRange('B15').setValue(''); //Brand
      SpreadsheetApp.getActiveSheet().getRange('B16').setValue(''); //Order #
      SpreadsheetApp.getActiveSheet().getRange('B13').setValue('USD'); //Currency
      SpreadsheetApp.getActiveSheet().getRange('B18').setValue('Score Autoreleased'); //Approver
      SpreadsheetApp.getActiveSheet().getRange('B19:B20').setValue('No'); //Verification | Represented
      SpreadsheetApp.getActiveSheet().getRange('B21:B26').setValue(''); //Rep Reason --> Documentation 4
      SpreadsheetApp.getActiveSheet().getRange('B17').setValue('0'); //Balance
      SpreadsheetApp.getActiveSheet().getRange('E3:F39').setValue(''); // Paste Area
      SpreadsheetApp.getActiveSheet().getRange('E3').setValue('Ctrl+SHIFT+V'); //Paste Cell
      SpreadsheetApp.getActiveSheet().getRange('B3').setFormula("=IF($E$4=\"\",\"\",IF($E$8=\"Transaction date:\",$F$8,IF(REGEXMATCH($E$3,\"Case ID\"),\"ENTER MANUALLY\",INDEX($G3:$G39,MATCH(\"PDT\",$I3:$I39,0)))))"); //Trans Date

      var lastRow = sheet2.getLastRow(); //Find last row of WorkLog
      sheet2.insertRowAfter(lastRow); //Append a blank row to the end of the WorkLog
  }

  function SendToWorkLogAD() { //This is the script for the Adyen sheet

      var sheet1 = ss.getSheetByName("AD");
      var rep = sheet1.getRange("B20");
      var repvalue = rep.getValue();
      var order = sheet1.getRange("B16");
      var ordervalue = order.getValue();
      var brand = sheet1.getRange("B15");
      var purchase = sheet1.getRange("B14");
      var approver = sheet1.getRange("B18");

      SpreadsheetApp.getActiveSheet().getRange('B27').setValue(user);

      if (ordervalue.indexOf('CNZ') === -1) {
          Browser.msgBox("Please enter a valid Order #");
          return
      } else if (brand.getValue() == "") {
          Browser.msgBox("Please make sure the 'Brand' field is populated");
          return
      } else if (brand.getValue() == "#N/A") {
          Browser.msgBox("This brand needs to be added to the Data Validation sheet");
          return
      } else if (purchase.getValue() == "") {
          Browser.msgBox("Please make sure the 'Purchase Amount' field is populated");
          return
      } else if (approver.getValue() == "") {
          Browser.msgBox("Please make sure the 'Approver' field is populated");
          return
      }

      if (repvalue == 'Yes') {

          sheet1.getRange("A2:Y2").copyTo(repsheet.getRange(repsheet.getLastRow() + 1, 1, 1, 7), {
              contentsOnly: true
          });
          var replastRow = repsheet.getLastRow(); //Find last row of RepSheet
          repsheet.insertRowAfter(replastRow); //Append a blank row to the end of the repsheet
      }

      sheet1.getRange("A2:Y2").copyTo(sheet2.getRange(sheet2.getLastRow() + 1, 1, 1, 7), {
          contentsOnly: true
      });
      SpreadsheetApp.getActiveSheet().getRange('B15').setValue(''); //Brand
      SpreadsheetApp.getActiveSheet().getRange('B9').setValue('Chargeback'); //Stage Reached
      SpreadsheetApp.getActiveSheet().getRange('B16').setValue(''); //Order #
      SpreadsheetApp.getActiveSheet().getRange('B18').setValue('Score Autoreleased'); //Approver
      SpreadsheetApp.getActiveSheet().getRange('B19:B20').setValue('No'); //Verification | Represented
      SpreadsheetApp.getActiveSheet().getRange('B21:B26').setValue(''); //Rep Reason --> Documentation 4
      SpreadsheetApp.getActiveSheet().getRange('B17').setValue('0'); //Balance
      SpreadsheetApp.getActiveSheet().getRange('E3:I8').setValue(''); //Paste area
      SpreadsheetApp.getActiveSheet().getRange('E3').setValue('Ctrl+SHIFT+V'); //Paste cell

      var lastRow = sheet2.getLastRow(); //Find last row of WorkLog
      sheet2.insertRowAfter(lastRow); //Append a blank row to the end of the WorkLog
  }
  
   function SendToWorkLogJCP() { //This is the script for the JCP sheet

      var sheet1 = ss.getSheetByName("JCP");
      var rep = sheet1.getRange("B20");
      var repvalue = rep.getValue();
      var order = sheet1.getRange("B16");
      var ordervalue = order.getValue();
      var brand = sheet1.getRange("B15");
      var purchase = sheet1.getRange("B14");
      var approver = sheet1.getRange("B18");

      SpreadsheetApp.getActiveSheet().getRange('B27').setValue(user);

      if (ordervalue.indexOf('CNZ') === -1) {
          Browser.msgBox("Please enter a valid Order #");
          return
      } else if (brand.getValue() == "") {
          Browser.msgBox("Please make sure the 'Brand' field is populated");
          return
      } else if (purchase.getValue() == "") {
          Browser.msgBox("Please make sure the 'Purchase Amount' field is populated");
          return
      } else if (approver.getValue() == "") {
          Browser.msgBox("Please make sure the 'Approver' field is populated");
          return
      }

      if (repvalue == 'Yes') {

          sheet1.getRange("A2:Y2").copyTo(repsheet.getRange(repsheet.getLastRow() + 1, 1, 1, 7), {
              contentsOnly: true
          });
          var replastRow = repsheet.getLastRow(); //Find last row of RepSheet
          repsheet.insertRowAfter(replastRow); //Append a blank row to the end of the repsheet
      }

      sheet1.getRange("A2:Y2").copyTo(sheet2.getRange(sheet2.getLastRow() + 1, 1, 1, 7), {
          contentsOnly: true
      });
      SpreadsheetApp.getActiveSheet().getRange('B9').setValue('Chargeback'); //Stage Reached
      SpreadsheetApp.getActiveSheet().getRange('B18').setValue('Score Autoreleased'); //Approver
      SpreadsheetApp.getActiveSheet().getRange('B19:B20').setValue('No'); //Verification | Represented
      SpreadsheetApp.getActiveSheet().getRange('B21:B26').setValue(''); //Rep Reason --> Documentation 4
      SpreadsheetApp.getActiveSheet().getRange('B17').setValue('0'); //Balance
      SpreadsheetApp.getActiveSheet().getRange('E3:U3').setValue(''); //Paste area
      SpreadsheetApp.getActiveSheet().getRange('E3').setValue('Ctrl+SHIFT+V'); //Paste cell

      var lastRow = sheet2.getLastRow(); //Find last row of WorkLog
      sheet2.insertRowAfter(lastRow); //Append a blank row to the end of the WorkLog
  }
