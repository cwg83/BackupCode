  function SendToWorkLogCL() { //This is the script for the ClientLine sheet

      var user = Session.getActiveUser().getUserLoginId();
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet2 = ss.getSheetByName("Work Log");
      var repsheet = ss.getSheetByName("Representing");
      var updates = ss.getSheetByName("Updates");      
      var sheet1 = ss.getSheetByName("CL");
      var rep = sheet1.getRange("B20");
      var repvalue = rep.getValue();
      var order = sheet1.getRange("B16");
      var ordervalue = order.getValue();
      var brand = sheet1.getRange("B11");
      var purchase = sheet1.getRange("B15");
      var stage = sheet1.getRange("B9");
      var approver = sheet1.getRange("B18");
      var formulacells = sheet1.getRange('B3:B15');
      var textcells = sheet1.getRange('B16:B26');
      var pastesection = sheet1.getRange('E3:F84');
      var pastecell = sheet1.getRange('E3');
      
      SpreadsheetApp.getActiveSheet().getRange('B27').setValue(user);

      if (ordervalue.indexOf('CNZ') === -1) {
          Browser.msgBox("Please enter a valid Order #");
          return
      } else if (brand.getValue() == "#N/A") {
          Browser.msgBox("This brand needs to be added to the Data Validation sheet");
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

      var formulas = [
          ['=IF($A$2="","",$E$17-1)'],
          ['=IF(LEFT($B$13,2)="RT",$E$22,$E$26)'],
          ['=IF($A$2="","",IF(LEFT($B$13,2)="RT",SUBSTITUTE($E$16,"  USD ",""),SUBSTITUTE($E$21,"  USD ","")))'],
          ['=IF(LEFT($B$13,2)="RT",$B$5,IF($E$20=0,$E$11,LEFT($E$20,LEN($E$20)-7)))'],
          ['=IF($A$2="","",TRIM($E$14))'],
          ['=IF(A2="","",TRIM(E8))'],
          ['=IF($A$2="","",IF(RegExMatch($E$25,"First Chargeback"),"Chargeback",IF(RegExMatch($E$25,"Second Chargeback"),"Second Chargeback",IF(RegExMatch($E$25,"Pre-Arbitration"),"Pre-Arbitration",IF(RegExMatch($E$26,"Retrieval"),"Retrieval","")))))'],
          ['=IF(A2="","",VLOOKUP(B8,\'Data Validation\'!T2:U,2,FALSE))'],
          ['=IF(A2="","",VLOOKUP($D$18,\'Data Validation\'!$C$2:$F,4,FALSE))'],
          ['=E33'],
          ['=IF($A$2="","",TRIM($E$5))'],
          ['=IF(A2="","",IF(REGEXMATCH($E$10,"CAD"),"CAD","USD"))'],
          ['=IF($B$14="USD",TRIM($B$5),"")']
      ];

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
          ['']
      ];

      formulacells.setFormulas(formulas);
      textcells.setValues(text);
      pastesection.setValue('');
      pastecell.setValue('Ctrl+SHIFT+V'); //Paste Cell
      var lastRow = sheet2.getLastRow();
      sheet2.insertRowAfter(lastRow); //Append a blank row to the end of the WorkLog
  }


  function SendToWorkLogCH() { //This is the script for the Chase sheet

      var user = Session.getActiveUser().getUserLoginId();
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet2 = ss.getSheetByName("Work Log");
      var repsheet = ss.getSheetByName("Representing");
      var updates = ss.getSheetByName("Updates");
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
      var formulacells = sheet1.getRange('B3:B15');
      var textcells = sheet1.getRange('B16:B26');
      var pastesection = sheet1.getRange('E3:F83');
      var pastecell = sheet1.getRange('E3');

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

      var formulas = [
          ['=IFERROR(LEFT(INDEX(F3:F84,MATCH(\"Auth Date\",E3:E84,0)),10),"")'],
          ['=IF(RegExMatch(E6,"Retrieval"),SUBSTITUTE(E7,"Case Status Date",""),IF(E18=" ",SUBSTITUTE(E7,"Case Status Date",""),SUBSTITUTE(E6,"Case Status Date","")))'],
          ['=IFERROR(LEFT(INDEX(F29:F84,MATCH("Amount (Presentment)",E29:E84,0)),LEN(INDEX(F29:F84,MATCH("Amount (Presentment)",E29:E84,0)))-5),"")'],
          ['=IFERROR(LEFT(INDEX(F3:F84,MATCH("Initial Chargeback Amount (Settlement)",E3:E84,0)),LEN(INDEX(F3:F84,MATCH("Initial Chargeback Amount (Settlement)",E3:E84,0)))-5),IFERROR(LEFT(INDEX(F3:F84,MATCH("Initial RR Amount (Presentment)",E3:E84,0)),LEN(INDEX(F3:F84,MATCH("Initial RR Amount (Presentment)",E3:E84,0)))-5),""))'],
          ['=IFERROR(INDEX(F3:F84,MATCH("Method of Payment",E3:E84,0)),"")'],
          ['=IF(E17="",SUBSTITUTE($E$19,"Reason Code #: ",""),SUBSTITUTE(SUBSTITUTE($E$20,"Reason Code #: ",""),"Reason Code: ",""))'],
          ['=IF(A2="","",IF(E17="",IF(RegExMatch($E$18,"Chargeback"),"Chargeback","Retrieval"),IF(RegExMatch($E$19,"Chargeback"),"Chargeback","Retrieval")))'],
          ['=IF(A2="","",VLOOKUP(B8,\'Data Validation\'!T2:U,2,FALSE))'],
          ['=IF(A2="","",IF(E17="",VLOOKUP($F$71,\'Data Validation\'!$E$2:$F,2,FALSE),VLOOKUP($F$72,\'Data Validation\'!$E$2:$F,2,FALSE)))'],
          ['=IF(RegExMatch(E6,"Retrieval"),F36,IF(E17="",$F$46,IF(E18="",F47,$F$45)))'],
          ['=IF(A2="","",E3)'],
          ['=IF(REGEXMATCH($F$26,"CAD"),"CAD",IF(REGEXMATCH($F$25,"CAD"),"CAD",IF(REGEXMATCH($F$27,"CAD"),"CAD","USD")))'],
          ['=IF(B14="USD",B5,SUBSTITUTE($F$27," (CAD)",""))']
      ];

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
          ['']
      ];

      formulacells.setFormulas(formulas);
      textcells.setValues(text);
      pastesection.setValue('');
      pastecell.setValue('Ctrl+SHIFT+V'); //Paste Cell
      var lastRow = sheet2.getLastRow();
      sheet2.insertRowAfter(lastRow); //Append a blank row to the end of the WorkLog
  }





  function SendToWorkLogAM() { //This is the script for the Amex sheet

      var user = Session.getActiveUser().getUserLoginId();
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet2 = ss.getSheetByName("Work Log");
      var repsheet = ss.getSheetByName("Representing");
      var updates = ss.getSheetByName("Updates");
      var sheet1 = ss.getSheetByName("AM");
      var rep = sheet1.getRange("B20");
      var repvalue = rep.getValue();
      var order = sheet1.getRange("B16");
      var ordervalue = order.getValue();
      var brand = sheet1.getRange("B11");
      var purchase = sheet1.getRange("B15");
      var approver = sheet1.getRange("B18");
      var formulacells = sheet1.getRange('B3:B15');
      var textcells = sheet1.getRange('B16:B26');
      var pastesection = sheet1.getRange('E3:F15');
      var pastecell = sheet1.getRange('E3');
      
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

      var formulas = [
          ['=IF(A4="","",IF(REGEXMATCH(E8,"CAD"),(RIGHT(E7,4)&"-"&MID(E7,4,2)&"-"&LEFT(E7,2)),E7))'],
          ['=IF(E4="","",IF(REGEXMATCH(E8,"CAD"),B29&"/"&A29&"/"&C29,E5))'],
          ['=IF(A2="","",IF(REGEXMATCH(E8,"CAD"),RIGHT(E8,LEN(E8)-3),RIGHT(E8,LEN(E8))))'],
          ['=IF(A2="","",IF(REGEXMATCH(E9,"CAD"),RIGHT(E9,LEN(E9)-3),IF(REGEXMATCH(E8,"CAD"),RIGHT(E8,LEN(E8)-3),IF(E9="0.00",SUBSTITUTE(E12,"-",""),E9))))'],
          ['=IF(A2="","","Amex")'],
          ['=IF(E6="Other","Not received",TRIM(E6))'],
          ['=IF($A$2="","",IF($E$13="IQ","Retrieval","Chargeback"))'],
          ['=IF(A2="","",VLOOKUP(B8,\'Data Validation\'!T2:U,2,FALSE))'],
          ['=IF(A2="","",VLOOKUP(D18,\'Data Validation\'!D2:F,3,FALSE))'],
          ['=E11'],
          ['=IF(A2="","",E3)'],
          ['=IF(REGEXMATCH(E8,"CAD"),"CAD","USD")'],
          ['=IF(A2="","",SUBSTITUTE(E8,"CAD",""))']
      ];

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
          ['']
      ];

      formulacells.setFormulas(formulas);
      textcells.setValues(text);
      pastesection.setValue('');
      pastecell.setValue('Ctrl+SHIFT+V'); //Paste Cell
      var lastRow = sheet2.getLastRow();
      sheet2.insertRowAfter(lastRow); //Append a blank row to the end of the WorkLog
  }


function SendToWorkLogAD() { //This is the script for the Adyen sheet

      var user = Session.getActiveUser().getUserLoginId();
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet2 = ss.getSheetByName("Work Log");
      var repsheet = ss.getSheetByName("Representing");
      var updates = ss.getSheetByName("Updates");
      var sheet1 = ss.getSheetByName("AD");
      var rep = sheet1.getRange("B20");
      var repvalue = rep.getValue();
      var order = sheet1.getRange("B16");
      var ordervalue = order.getValue();
      var brand = sheet1.getRange("B15");
      var purchase = sheet1.getRange("B14");
      var approver = sheet1.getRange("B18");
      var formulacells = sheet1.getRange('B3:B14');
      var textcells = sheet1.getRange('B15:B26');
      var pastesection = sheet1.getRange('E3:I8');
      var pastecell = sheet1.getRange('E3');
      
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
      var formulas = [
          ['=IF(E4="","",LEFT(I4,LEN(I4)-13))'],
          ['=IF(E4="","",DATEVALUE(LEFT(I3,LEN(I3)-13)))'],
          ['=IF(A2="","",RIGHT(I6,LEN(I6)-4))'],
          ['=IF(A2="","",RIGHT(I5,LEN(I5)-4))'],
          ['=IF(A2="","",F6)'],
          ['=TRIM(F3)'],
          ['=IF(A2="","","Chargeback")'],
          ['=IF(A2="","",VLOOKUP(B8,\'Data Validation\'!T2:U,2,FALSE))'],
          ['=I8'],
          ['=IF(A2="","",F8)'],
          ['=IF(A2="","",IF(RegExMatch(I6,"GBP"),"GBP",IF(RegExMatch(I6,"HKD"),"HKD",IF(RegExMatch(I6,"EUR"),"EUR"))))'],
          ['=IF($A$2="","",$B$5)'],
       
      ];

      var text = [
          [''],
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
          ['']
      ];

      formulacells.setFormulas(formulas);
      textcells.setValues(text);
      pastesection.setValue('');
      pastecell.setValue('Ctrl+SHIFT+V'); //Paste Cell
      var lastRow = sheet2.getLastRow();
      sheet2.insertRowAfter(lastRow); //Append a blank row to the end of the WorkLog
  }


  function SendToWorkLogJCP() { //This is the script for the JCP sheet

      var user = Session.getActiveUser().getUserLoginId();
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet2 = ss.getSheetByName("Work Log");
      var repsheet = ss.getSheetByName("Representing");
      var updates = ss.getSheetByName("Updates");
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

  function SendToWorkLogPP() { //This is the script for the PayPal sheet
  
      var user = Session.getActiveUser().getUserLoginId();
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet2 = ss.getSheetByName("Work Log");
      var repsheet = ss.getSheetByName("Representing");
      var updates = ss.getSheetByName("Updates");
      var sheet1 = ss.getSheetByName("PP");
      var rep = sheet1.getRange("B20");
      var repvalue = rep.getValue();
      var order = sheet1.getRange("B16");
      var ordervalue = order.getValue();
      var brand = sheet1.getRange("B15");
      var purchase = sheet1.getRange("B14");
      var approver = sheet1.getRange("B18");
      var formulacells = sheet1.getRange('B3:B14');
      var textcells = sheet1.getRange('B15:B26');
      var pastesection = sheet1.getRange('E3:F39');
      var pastecell = sheet1.getRange('E3');
      
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

      var formulas = [
          ['=IF($E$3="Transaction ID:",INDEX($F$3:$F$17,MATCH("Transaction Date:",$E$3:$E$17,0)),IF(REGEXMATCH($E$3,"PP-D"),"ENTER MANUALLY",""))'],
          ['=IF(E2="","",IF(REGEXMATCH($E$3,"PP-D"),"ENTER MANUALLY",IFERROR(INDEX($F$3:$F$17,MATCH("Date of Complaint:",$E$3:$E$17,0)),INDEX($F$3:$F$17,MATCH("Chargeback Date:",$E$3:$E$17,0)))))'],
          ['=IF(E2="","",IFERROR(SUBSTITUTE(INDEX($F$3:$F$17,MATCH("Transaction Amount:",$E$3:$E$17,0))," USD",""),INDEX($F$3:$F$17,MATCH("Transaction Amount",$E$3:$E$17,0))))'],
          ['=IFERROR(IF($E$3="Transaction ID:",SUBSTITUTE(INDEX($F$3:$F$17,MATCH("Chargeback Amount:",$E$3:$E$17,0))," USD",""),SUBSTITUTE(INDEX($F$3:$F$17,MATCH("Disputed Amount:",$E$3:$E$17,0))," USD")),$B$5)'],
          ['=IF(E2="","","PayPal")'],
          ['=IF(E2="","",IFERROR(IF($E$3="Transaction ID:",INDEX($F$3:$F$17,MATCH("Reason for Dispute:",$E$3:$E$17,0)),INDEX($F$3:$F$17,MATCH("Dispute reason:",$E$3:$E$17,0))),INDEX($F$3:$F$17,MATCH("Dispute reason",$E$3:$E$17,0))))'],
          ['=IF(E2="","",IF($B$8="Unauthorized Payment","Chargeback",IF($B$8="Merchandise","Chargeback",IF($B$8="Item not received","Chargeback","Retrieval"))))'],
          ['=IFERROR(VLOOKUP($B$8,\'Data Validation\'!$T$2:U,2,FALSE),"")'],
          ['=IF(E2="","",IFERROR(INDEX($F$3:$F$17,MATCH("Invoice ID:",$E$3:$E$17,0)),INDEX($F$3:$F$17,MATCH("Invoice ID",$E$3:$E$17,0))))'],
          ['=IF($E$3="Transaction ID:",SUBSTITUTE(INDEX($F$3:$F$17,MATCH("PayPal Case ID:",$E$3:$E$17,0)),"PP-",""),IF(REGEXMATCH($E$3,"PP-D"),SUBSTITUTE($E$3,"Case ID: ",""),""))'],
          ['=IF(E2="","","USD")'],
          ['=$B$5'],          
       
      ];

      var text = [
          [''],
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
          ['']
      ];

      formulacells.setFormulas(formulas);
      textcells.setValues(text);
      pastesection.setValue('');
      pastecell.setValue('Ctrl+SHIFT+V'); //Paste Cell
      var lastRow = sheet2.getLastRow();
      sheet2.insertRowAfter(lastRow); //Append a blank row to the end of the WorkLog
  }


