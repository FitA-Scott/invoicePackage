// Global Variables
  var ui = SpreadsheetApp.getUi();
  var active = SpreadsheetApp.getActiveSpreadsheet();
  var detail = active.getSheetByName('Invoice');

function onOpen(e) {
    ui.createMenu('Billing')
      .addItem('Show Navigation Panel', 'showSidebar')
      .addItem('Refresh Data', 'importCustomerData')
      .addItem('Refresh Line Items','assembleLineItems')
      .addItem('Send Proof for Approval','approvalProcess')
      .addItem('Open Contracts Folder','viewContract')
      .addItem('Open Invoices Folder','viewInvoices')
      .addToUi();     
      searchNumber();
      importCustomerData();
      pullBillingInfo();
  var testRange = active.getSheetByName('Details').getRange(23,2,1,1).getValue();
    if ( testRange != '') {
      specialSalesData();
  }
} 

function showSidebar() {
  var list = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Invoice Functions')
      .setWidth(300);
  SpreadsheetApp.getUi()
      .showSidebar(list);
}
function routeProcess(){
  var reviewTest = SpreadsheetApp.getActive().getSheetByName('Calculations').getRange(3,2,1,1).getValue();
  var approvalTest = SpreadsheetApp.getActive().getSheetByName('Calculations').getRange(4,2,1,1).getValue();
      if (reviewTest == 'Not Required'){
      savePDF();
      }
      if (reviewTest == 'Required'){
        if (approvalTest == 'Approved'){
        savePDF();
        }      
        if (approvalTest != 'Approved'){
        approvalProcess();
        }
      }   
}

function approvalProcess( optSSId, optSheetId ){
  var workingSheet = SpreadsheetApp.getActive();
  var calcSheet = workingSheet.getSheetByName('Details');
  var invoiceSheet = workingSheet.getSheetByName('Invoice');
  var companyName = calcSheet.getRange(3,2,1,1).getValue();
  var invoiceNumber = invoiceSheet.getRange(9,7,1,1).getValue();
  var invoicePeriod = invoiceSheet.getRange(10,7,1,1).getValue();
  var billingDocId = workingSheet.getId();
  var accountNumber = calcSheet.getRange(5,2,1,1).getValue();
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
  'Approval is required for this invoice. Enter a short message for the approver.',
  ui.ButtonSet.OK_CANCEL);
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button = 'OK'){
  var messageSubject = '[Invoice Review] '+ companyName +  ' for the service period ' + invoicePeriod;
  var linkToForm= 'https://docs.google.com/forms/d/e/1FAIpQLSekfBkeUAYiFwMMiKtZBoVcuqRorYOtHqfRpE9QAEdHwvFsVQ/viewform?usp=pp_url&entry.280149682='+companyName+'&entry.163194497='+invoiceNumber+'&entry.974011626='+billingDocId+'&entry.1468691854='+accountNumber;
  var htmlButton = '<table width="100%" cellspacing="0" cellpadding="0"><tr><td><table cellspacing="0" cellpadding="0"><tr><td style="border-radius: 4px;" bgcolor=“#34495E”><a href="'+ linkToForm +'" target="_blank" style="padding: 8px 12px; border: 1px solid #34495E;border-radius: 4px;font-family: Helvetica, Arial, sans-serif;font-size: 14px; color: #ffffff;text-decoration: none;font-weight:bold;display: inline-block;">Go To Response Form</a></td></tr></table></td></tr></table>';
  var messageBody = 'Hi All,<p><p>An invoice was created for '+companyName+' that was flagged for review in the Billing Summaries sheet. <p><p>' + 'The Billing team has included the following message:<p><p>' + text + '<p><p>A copy of this invoice can be found attached to this email. Could you please review this invoice and provide an approval or rejection response using the form found by clicking the button below.<br><br><br>'+htmlButton+'<br><br><br>Kind Regards,<p><p><p>Finance and Legal Team';
  var ss = (optSSId) ? SpreadsheetApp.openById(optSSId) : SpreadsheetApp.getActiveSpreadsheet();
  var url = ss.getUrl().replace(/edit$/,'');
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  if (parents.hasNext()) {
    var folder = parents.next();
  }
  else {
    folder = DriveApp.getRootFolder();
  }
  var sheets = ss.getSheets();
  var slice = sheets.slice(0,1);
  for (var i=0; i<slice.length; i++) {
    var sheet = slice[i];
    if (optSheetId && optSheetId !== sheet.getSheetId() )continue; 
    var url_ext = 'export?exportFormat=pdf&format=pdf'
        + '&gid=' + sheet.getSheetId()
        // Optional for PDF
        + '&size=A4'
        + '&portrait=true'
        + '&fitw=true'
        + '&sheetnames=false&printtitle=false&pagenumbers=false'
        + '&gridlines=false'
        + '&fzr=false';
    var header = {
      headers: {
        'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()
      }
    }
    var response = UrlFetchApp.fetch(url + url_ext, header);
    var blob = response.getBlob().setName('[INVOICE PROOF]_' + invoiceNumber + '_' + companyName + '_' + invoicePeriod + '.pdf');
    folder.createFile(blob);
  var approvers = 'schulze@fitanalytics.com; shenhav@fitanalytics.com';
    GmailApp.sendEmail(approvers,messageSubject,'',{name:'Accounts Receivable',from:'invoices@fitanalytics.com',replyto:'invoices@fitanalytics.com', htmlBody: messageBody, attachments:[blob.getAs(MimeType.PDF)]}); 
  var requestSheet = SpreadsheetApp.openById('1s2RFxnJTBIfIWg64gGqhYv95v0wWTXN0Qp4H4_NvWFA').getSheetByName('Requests');
  var requestNumberRange = requestSheet.getRange(requestSheet.getLastRow()+1,1,1,1);
  var requestNameRange = requestSheet.getRange(requestSheet.getLastRow()+1,2,1,1);
  var requesterId = Session.getEffectiveUser();
  var requesterRange = requestSheet.getRange(requestSheet.getLastRow()+1,3,1,1);
  var requestDate = new Date();
  var requestDateRange = requestSheet.getRange(requestSheet.getLastRow()+1,4,1,1);
  requestNumberRange.setValue(invoiceNumber);
  requestNameRange.setValue(companyName);
  requesterRange.setValue(requesterId);
  requestDateRange.setValue(requestDate);
    }
  }
}

function createCancellation(){
  var cancellation = HtmlService.createHtmlOutputFromFile('Details')
      .setWidth(450)
      .setHeight(200);    
  SpreadsheetApp.getUi().showModalDialog(cancellation, 'Enter Invoice Details');
}

function logData(billmonth, billyear) {
  var detailsheet = SpreadsheetApp.getActive()
  var calc = detailsheet.getSheetByName('Calculations');
  var history = detailsheet.getSheetByName('Historical Data');
  var billinglog = SpreadsheetApp.openById('1D5VqWLYIk3FiDHEyFmqn8XDwOerrQZEOg1hKHnoH6aw');
  var numbersheet = billinglog.getSheetByName('Invoice Numbers');
  var logsheet = billinglog.getSheetByName('Billing Log');
  var newcanc = logsheet.getRange(logsheet.getLastRow()+1,2,1,1);
  var oldnum = calc.getRange(11,2,1,1).getValue();
  var newnum = numbersheet.getRange(2,2,1,1).getValue();
  var invoicetype = 'Cancellation';
  var typeRange = calc.getRange(8,2,1,1);
  var multiplierRange = calc.getRange(12,2,1,1);
  var multiplier = '-1';
  var newinvnum = calc.getRange(16,5,1,1);
    typeRange.setValue(invoicetype);  
    multiplierRange.setValue(multiplier);
    newinvnum.setValue(newnum);
    newcanc.setValue(oldnum);
}
function cancelInvoice(){
  var detailsheet = SpreadsheetApp.getActive()
  var calc = detailsheet.getSheetByName('Calculations');
  var history = detailsheet.getSheetByName('Historical Data');
  var billinglog = SpreadsheetApp.openById('1D5VqWLYIk3FiDHEyFmqn8XDwOerrQZEOg1hKHnoH6aw');
  var logsheet = billinglog.getSheetByName('Billing Log');
  var newrow = logsheet.getRange(logsheet.getLastRow()+1,2,1,1);
  var newname = logsheet.getRange(logsheet.getLastRow()+1,3,1,1);
  var prefix = calc.getRange(29,1,1,1).getValue();
  var newnum = calc.getRange(16,5,1,1).getValue();
  var status = 'Cancelled';
  var invnum = calc.getRange(9,2,1,1).getValue(); 
  var rowFinder = history.createTextFinder(invnum);
  var rowNum = rowFinder.findNext().getRow();
  var newstatus = history.getRange(rowNum,6,1,1);
  newrow.setValue(newnum);
  newname.setValue(prefix);
  newstatus.setValue(status);
  buildHistoricalLineItem();
}

function savePDF( optSSId, optSheetId ) {
  var calculationSource = SpreadsheetApp.getActive().getSheetByName('Details');
  var invoiceSource = SpreadsheetApp.getActive().getSheetByName('Invoice');
  var invoiceNumber = invoiceSource.getRange(9,7,1,1).getValue();
  var accountNameText = calculationSource.getRange(3,2,1,1).getValue();
  var accountName = accountNameText.replace(" ","_",'g');
  var servicePeriodText = invoiceSource.getRange(10,7,1,1).getValue();
  var servicePeriod = servicePeriodText.replace(" ","_");
  var ss = (optSSId) ? SpreadsheetApp.openById(optSSId) : SpreadsheetApp.getActiveSpreadsheet();
  var url = ss.getUrl().replace(/edit$/,'');
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  if (parents.hasNext()) {
    var folder = parents.next();
  }
  else {
    folder = DriveApp.getRootFolder();
  }
  var sheets = ss.getSheets();
  var slice = sheets.slice(0,1);
  for (var i=0; i<slice.length; i++) {
    var sheet = slice[i];
    if (optSheetId && optSheetId !== sheet.getSheetId() )continue; 
    var url_ext = 'export?exportFormat=pdf&format=pdf'
        + '&gid=' + sheet.getSheetId()
        // Optional for PDF
        + '&size=A4'
        + '&portrait=true'
        + '&fitw=true'
        + '&sheetnames=false&printtitle=false&pagenumbers=false'
        + '&gridlines=false'
        + '&fzr=false';
    var header = {
      headers: {
        'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()
      }
    }
    var response = UrlFetchApp.fetch(url + url_ext, header);
    var blob = response.getBlob().setName('Fit_Analytics_' + invoiceNumber + '_' + accountName + '_' + servicePeriod + '.pdf');
    folder.createFile(blob);
  }
    var result = ui.alert('Would you like to e-mail the invoice?', ui.ButtonSet.YES_NO_CANCEL);
  // Process user response.
  if (result == ui.Button.YES) {
    ui.alert('An e-mail draft has been created, please proofread and send.');
  //Prepare draft to send to client and second email to send to Digi-Bel and Salesforce
  var infosheet = SpreadsheetApp.getActive();
  var infosource = infosheet.getSheetByName('Invoice');
  var detailsource = infosheet.getSheetByName('Details');
  var calcsource = infosheet.getSheetByName('Calculations');  
  var companyname = detailsource.getRange(3,2,1,1).getValues();
  var invoiceperiod = infosource.getRange(9,7,1,1).getValues();
  var bccaddress = calcsource.getRange(17,2,1,1).getValues();
  var deliveryaddresses = detailsource.getRange(17,2,1,1).getValues();
  var emailsubject = calcsource.getRange(22,2,1,1).getValues();
  var sheetbodytext = calcsource.getRange(23,2,1,1).getValues();
  var emailtext = String(sheetbodytext[0]);
  var sfdcid= detailsource.getRange(21,2,1,1).getValue();
  var reviewTest = calcsource.getRange(2,2,1,1).getValue();
  var approvalTest = calcsource.getRange(4,2,1,1).getValue();
  var emailfooter = ('<div><br><br><br><br><br>Kind regards</div><br><b>Fit Analytics Accounting Team</b><br><br><img src="https://ci5.googleusercontent.com/proxy/92ywHWBtnnjrrcbYhVDoqWjHZNDKD2ukCvaIDfIoFxERJKyIfwLaSW13NVs2ECuVzo63kHv6ZIpZMuPWjBlr28gADggLhp-h4p5qhcQ37au1-aDY2xQTaB9sOGNKtkGk3Rvs5Ze8Xv4C4rjPmYfSrp__0mwmpG5q0THAh84N8eiA3K1HnYXb4OnvuZC4IOZKlJXTDZs64C8=s0-d-e1-ft#https://docs.google.com/uc?export=download&amp;id=0B0gpnzRVY698NUN3WGJoWEk1NXc&amp;revid=0B0gpnzRVY698aUFoUitYeDNpQTRCNWtqTW9VWEtkbGlmK2lJPQ" alt="" width="169" height="40" style="font-family:arial,helvetica,sans-serif;font-size:12.8px" class="CToWUd"></div><div style="font-size:11.1px; color:#666666" ><b>SOLVE SIZING. SELL SMARTER.<b></div><br><div>Voigtstraße 3 | 10247 Berlin</div><br><div>www.fitanalytics.com</div>');
  GmailApp.createDraft(deliveryaddresses, emailsubject,'',{ name: 'Fit Analytics GmbH Accounts Receivable', from: 'invoices@fitanalytics.com', replyto: 'invoices@fitanalytics.com', htmlBody: emailtext + emailfooter, bcc: 'invoices@fitanalytics.com; puz.7002@digi-bel.de', attachments:[blob.getAs(MimeType.PDF)]});  
  MailApp.sendEmail('emailtosalesforce@18xzv579vg9bl3mjpl6uzyy6ho177oxejfjuyovc7o6jozgn53.0o-s6v5uai.eu9.le.salesforce.com','[Invoice] for ' + companyname + 'for ' + invoiceperiod, 'ref: ' + sfdcid, { name: 'General FitA', attachments:[blob.getAs(MimeType.PDF)]}); 
  moveBillingLogLineItem()
  var cancellationtest = calcsource.getRange(8,2,1,1).getValue();
     if (cancellationtest == "Cancellation"){
       cancelInvoice();
     }
  // Process alternate user response  
  } else if (result == ui.Button.NO) {
    ui.alert('Print invoice before closing. Please note that a copy of the email has been sent to both Digi-Bel and Salesforce.');
    MailApp.sendEmail('emailtosalesforce@18xzv579vg9bl3mjpl6uzyy6ho177oxejfjuyovc7o6jozgn53.0o-s6v5uai.eu9.le.salesforce.com, puz.7002@digi-bel.de','[Invoice] for ' + companyname + 'for ' + invoiceperiod, 'ref: ' + sfdcid, { name: 'General FitA', attachments:[blob.getAs(MimeType.PDF)]}); 
    moveBillingLogLineItem()
     if (cancellationtest == "Cancellation"){
       cancelInvoice();
     }
  // User cancels process
  } else if (result == ui.Button.CANCEL) {
    ui.alert('Process cancelled, the invoice has been created but not sent');
  // User closes wndow with no response
  } else if (result == ui.Button.CLOSE) {
    ui.alert('Process cancelled, the invoice has been created but not sent');
  }
}
function moveBillingLogLineItem() {
  var destinationSheet = SpreadsheetApp.openById('1D5VqWLYIk3FiDHEyFmqn8XDwOerrQZEOg1hKHnoH6aw').getSheetByName('Incoming Line Items');
  var destinationRange = destinationSheet.getRange(destinationSheet.getLastRow()+1,1,1,15);
  var cancellationSheet = SpreadsheetApp.openById('1D5VqWLYIk3FiDHEyFmqn8XDwOerrQZEOg1hKHnoH6aw').getSheetByName('Cancellations');
  var cancellationRange = cancellationSheet.getRange(cancellationSheet.getLastRow()+1,1,1,15);
  var calcsheet = SpreadsheetApp.getActive().getSheetByName('Calculations');
  var billingLogLineItem = calcsheet.getRange(26,1,1,15).getValues();
  var test = calcsheet.getRange(8,2,1,1).getValue();
  if (test == 'Regular'){
      destinationRange.setValues(billingLogLineItem);
  }
  else if (test == 'Cancellation') {
      cancellationRange.setValues(billingLogLineItem);
      cancelInvoice();
      resetSheet();
  }
}

function importCustomerData() {
  var activeSheet = SpreadsheetApp.getActive().getSheetByName('Details');
  var detailSheet = SpreadsheetApp.openById('1WQBEVDTyK8XvTG5BkMJMbqWMyKTf3aYuFjCQPuc23GI').getSheetByName('Client Info');  
  var updateRowNum = searchNumber(); 
  var updateInfo = detailSheet.getRange(updateRowNum,1,1,27).getValues();
  var updateSheet = SpreadsheetApp.getActive().getSheetByName('Customer Data');
  var updateDest = updateSheet.getRange(updateSheet.getLastRow()+1,1,1,27);
  updateDest.setValues(updateInfo);
  refreshCustomerData();
}

function searchNumber() {
  var originSheet = SpreadsheetApp.getActive().getSheetByName('Details');
  var valuesSheet = SpreadsheetApp.openById('1WQBEVDTyK8XvTG5BkMJMbqWMyKTf3aYuFjCQPuc23GI').getSheetByName('Client Info');  
  var accountNumber = originSheet.getRange(5,2,1,1).getValue();
  var updateValues = valuesSheet.getDataRange().getValues();
     for (var i = 0; i < updateValues.length; i++){
     for (var j = 0; j < updateValues[i].length; j++){
        if(updateValues[i][j] == accountNumber){
    return i+1;
        }
      }
    }  
  }

function specialSalesData() {
  var itemsheet = SpreadsheetApp.getActive().getSheetByName('Itemised Info');
  var detsheet = SpreadsheetApp.getActive().getSheetByName('Details');
  var calcsheet = SpreadsheetApp.getActive().getSheetByName('Calculations');
  var currentmonth = calcsheet.getRange(6,5,1,1).getValue();
  var datalocation = calcsheet.getRange(23,2,1,1).getValue();
  var purchaseFormula = '=IMPORTRANGE("https://docs.google.com/spreadsheets/d/' + datalocation + '", "Purchases ' + currentmonth + '!A1:J")';
  var returnFormula = '=IMPORTRANGE("https://docs.google.com/spreadsheets/d/' + datalocation + '", "Returns ' + currentmonth + '!A1:I")';
  var purchaseFormulaDest = itemsheet.getRange(1,1,1,1);
  var returnFormulaDest = itemsheet.getRange(1,12,1,1);
    purchaseFormulaDest.setValue(purchaseFormula);
    returnFormulaDest.setValue(returnFormula);
    setSpecialData();  
}

function setSpecialData(){
  var datasheet = SpreadsheetApp.getActive().getSheetByName('Itemised Info');
  var data = datasheet.getRange(1,1,datasheet.getLastRow(),datasheet.getLastColumn());
  var range = data.getA1Notation();
  var copy = data.getValues();
      datasheet.clear({contentsOnly: true});
      datasheet.getRange(range).setValues(copy);
}

function refreshCustomerData() {
  var current = SpreadsheetApp.getActive();
  var inbound = current.getSheetByName('Customer Data');
  var destination = current.getSheetByName('Details');
  var refreshDate = new Date();
  var newRefreshDate = inbound.getRange(inbound.getLastRow(),21,1,1);
  var userName = Session.getEffectiveUser();
  var newUserName = inbound.getRange(inbound.getLastRow(),22,1,1);
  //New values imported form Salesforce report
  var commonName = inbound.getRange(inbound.getLastRow(),3,1,1).getValue();
  var prefix = inbound.getRange(inbound.getLastRow(),4,1,1).getValue();
  var legalName = inbound.getRange(inbound.getLastRow(),5,1,1).getValue();
  var addressOne = inbound.getRange(inbound.getLastRow(),6,1,1).getValue();
  var addressTwo = inbound.getRange(inbound.getLastRow(),7,1,1).getValue();
  var addressThree = inbound.getRange(inbound.getLastRow(),8,1,1).getValue();
  var addressFour = inbound.getRange(inbound.getLastRow(),9,1,1).getValue();
  var billingContacts = inbound.getRange(inbound.getLastRow(),10,1,1).getValue();
  var billingEmails = inbound.getRange(inbound.getLastRow(),11,1,1).getValue();
  var currency = inbound.getRange(inbound.getLastRow(),12,1,1).getValue(); 
  var billMethod = inbound.getRange(inbound.getLastRow(),13,1,1).getValue();
  var billcycle = inbound.getRange(inbound.getLastRow(),14,1,1).getValue();
  var vatId = inbound.getRange(inbound.getLastRow(),15,1,1).getValue();
  var poNumber = inbound.getRange(inbound.getLastRow(),16,1,1).getValue();
  var paymentTerms = inbound.getRange(inbound.getLastRow(),17,1,1).getValue();
  var costCenter = inbound.getRange(inbound.getLastRow(),18,1,1).getValue();
  var salesforceId = inbound.getRange(inbound.getLastRow(),19,1,1).getValue();
  var contractFolder = inbound.getRange(inbound.getLastRow(),20,1,1).getValue();
  // Destinations for the new values
  var newCommonName = destination.getRange(2,2,1,1);
  var newLegalName = destination.getRange(3,2,1,1);
  var newPrefix = destination.getRange(4,2,1,1);
  var newAddressOne = destination.getRange(6,2,1,1);
  var newAddressTwo = destination.getRange(7,2,1,1);
  var newAddressThree = destination.getRange(8,2,1,1);
  var newAddressFour = destination.getRange(9,2,1,1);
  var newBillingContacts = destination.getRange(15,2,1,1);
  var newBillingEmails = destination.getRange(17,2,1,1);
  var newCurrency = destination.getRange(18,2,1,1);
  var newBillMethod = destination.getRange(20,2,1,1);
  var newBillCycle = destination.getRange(19,2,1,1);
  var newVatId = destination.getRange(10,2,1,1);
  var newPoNumber = destination.getRange(12,2,1,1);
  var newPaymentTerms = destination.getRange(14,2,1,1);
  var newCostCenter = destination.getRange(13,2,1,1);
  var newsalesforceId = destination.getRange(21,2,1,1);
  var newContractFolder = destination.getRange(22,2,1,1);
  // Constant Values that should be moved every time the refresh is run
  newCommonName.setValue(commonName);
  newLegalName.setValue(legalName);
  newPrefix.setValue(prefix);
  newAddressOne.setValue(addressOne);
  newAddressTwo.setValue(addressTwo);
  newAddressThree.setValue(addressThree);
  newAddressFour.setValue(addressFour);
  newBillingEmails.setValue(billingEmails);  
  newCurrency.setValue(currency);
  newBillMethod.setValue(billMethod);
  newBillCycle.setValue(billcycle);
  newVatId.setValue(vatId);
  newPoNumber.setValue(poNumber);
  newPaymentTerms.setValue(paymentTerms);
  newCostCenter.setValue(costCenter);
  newsalesforceId.setValue(salesforceId);
  newContractFolder.setValue(contractFolder);  
  newRefreshDate.setValue(refreshDate);
  newUserName.setValue(userName);  
  if (billingContacts == '') {
    newBillingContacts.setValue('Accounts Payable');
    }
    else {
    newBillingContacts.setValue(billingContacts);
    }
}

function CopyTemplate() {
  var newui = SpreadsheetApp.getUi();
  var newprompt = newui.prompt(
  'Enter password',
  newui.ButtonSet.OK_CANCEL);
  var click = newprompt.getSelectedButton();
  var password = newprompt.getResponseText();
  if (click == newui.Button.OK && password == 'masterclone') {        
  var active = SpreadsheetApp.getActive();
  var tab = active.getSheetByName('Details');
  var legalname = tab.getRange(3,2,1,1).getValue();
  var clientnumber = tab.getRange(5,2,1,1).getValue();
  var filename = (legalname + "-" + clientnumber + "-Invoicing File");
  var destfolder = DriveApp.getFolderById('1aAQef0Op-BEfjq2F2WKpzn7sf_4hEbUc');
  var newdoc = DriveApp.getFileById(active.getId()).makeCopy(filename, destfolder)
  var newdocid = newdoc.getId();
        destfolder.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  var newdoc = destfolder.get
  var url = "https://docs.google.com/spreadsheets/d/"+newdocid;
  var openNew = "<script>window.open('" + url + "');google.script.host.close();</script>";
  var userInterface = HtmlService.createHtmlOutput(openNew);
        newui.showModalDialog(userInterface, "Opening New Invoice File");        
    }
  else if (click == newui.Button.OK && password != 'masterclone') {
    newui.alert('Password incorrect.'); }  
  else if (click == newui.Button.CANCEL) {}
  else if (click == newui.Button.CLOSE) {}
}

function requirePassword(){
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
  'Enter password',
  ui.ButtonSet.OK_CANCEL);
  
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK && text == 'masterupdate' && SpreadsheetApp.getActive().getId() != 'M6KI4FZZCbq3pMjhKpJGSpUVA7j-W5aRr') {
      var sheet = SpreadsheetApp.getActive();
      var tab = sheet.getSheetByName('Details');
      var newdocid = sheet.getId();
      var companyname = tab.getRange(2,2,1,1).getValue();
      var clientnumber = tab.getRange(5,2,1,1).getValue();
      var targetsheet = SpreadsheetApp.openById('1WQBEVDTyK8XvTG5BkMJMbqWMyKTf3aYuFjCQPuc23GI');
      var targettab = targetsheet.getSheetByName('Client Master List');
      var targettabdata = targettab.getDataRange();
      var newrow = targettabdata.getLastRow()+1;
      var targetcompanyname = targettab.getRange(newrow,2,1,1);
      var targetclientnumber = targettab.getRange(newrow,1,1,1);
      var targetdocid = targettab.getRange(newrow,4,1,1);
      var commonname = tab.getRange(2,2,1,1).getValue();
      var urlformula = '=HYPERLINK(CONCATENATE("https://docs.google.com/spreadsheets/d/",RC[-1],"/edit#gid=712059032"),RC[-3])';
      var targetcommonname = targettab.getRange(newrow,3,1,1);
      var targeturlformula = targettab.getRange(newrow,5,1,1);
  
      targetcompanyname.setValue(companyname);
      targetclientnumber.setValue(clientnumber);
      targetdocid.setValue(newdocid);
      targetcommonname.setValue(commonname);
      targeturlformula.setValue(urlformula);
    }
  else if (button == ui.Button.OK && text != 'masterupdate') {
    ui.alert('Password incorrect.'); }
    else if (button == ui.Button.OK && SpreadsheetApp.getActive().getId() == 'M6KI4FZZCbq3pMjhKpJGSpUVA7j-W5aRr') {
    ui.alert('You cannot add the Template to the Master List.'); }  
  else if (button == ui.Button.CANCEL) {}
  else if (button == ui.Button.CLOSE) {}  
}
function viewContract() {
  var folder = SpreadsheetApp.getActive().getSheetByName('Details').getRange(22,2,1,1).getValue();
  var newModal = '<script>window.open("https://drive.google.com/drive/folders/' + folder + '");google.script.host.close();</script>';
  var interface = HtmlService.createHtmlOutput(newModal)
  SpreadsheetApp.getUi().showModalDialog(interface, 'Opening contracts folder');
}

function resetSheet() {
 var document = SpreadsheetApp.getActive();
 var calculations = document.getSheetByName('Calculations');
 var approvalRangeOne = calculations.getRange(4,2,4,1);
 var cancellationRangeOne = calculations.getRange(8,2,1,1);
 var clearString = '';
 var regularString = 'Regular';
 var multiplyRange = calculations.getRange(12,2,1,1);
 var multiply = '1'; 
    approvalRangeOne.setValue(clearString);
    cancellationRangeOne.setValue(regularString);
    multiplyRange.setValue(multiply);
}
function openAdminPanel(){
    var panel = HtmlService.createHtmlOutputFromFile('AdminPanel')
      .setTitle('Admin Panel')
      .setWidth(300);
  SpreadsheetApp.getUi()
      .showSidebar(panel);
}

function pullBillingInfo() {
  var infosheet = SpreadsheetApp.getActive().getSheetByName('Purchase Data');
  var detailsheet = SpreadsheetApp.getActive().getSheetByName('Calculations');
  var itemssheet = SpreadsheetApp.getActive().getSheetByName('Line Items');
  var purloc = detailsheet.getRange(7,5,1,1).getValue();
  var retloc = detailsheet.getRange(8,5,1,1).getValue();
  var purformula = '=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1n0oFePjP3SGpE9fK2j_IptZ93RySbGbZB12T7wz9Bhg/","' + purloc + '!A1:G")';
  var retformula = '=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1n0oFePjP3SGpE9fK2j_IptZ93RySbGbZB12T7wz9Bhg/","' + retloc + '!A1:G")'; 
  var purchases = infosheet.getRange(1,1,1,1);
  var returns = infosheet.getRange(1,9,1,1);
  var formula = '=IMPORTRANGE("1D7HfOkKW7k752Abclg2Aam65dlRYFyMxYJ2IDBfDkGE","List!A1:N")';
  var range = itemssheet.getRange(1,1,1,1);
    range.setValue(formula);
    purchases.setValue(purformula);
    returns.setValue(retformula);
    assembleLineItems();
}

function createOneOff(){
  var oneoff = HtmlService.createHtmlOutputFromFile('OneOffInvoice')
      .setWidth(450)
      .setHeight(400);    
  SpreadsheetApp.getUi().showModalDialog(oneoff, 'Enter Invoice Details');
}

function assembleLineItems(){
  var sheet = SpreadsheetApp.getActive();
  var items = sheet.getSheetByName('Line Items');
  var details = sheet.getSheetByName('Details');
  var calculations = sheet.getSheetByName('Calculations');  
  var quantity = calculations.getRange(5,5,1,1).getValue();
  var prefix = details.getRange(4,2,1,1).getValue();
  var prefixFinder = items.createTextFinder(prefix);
  var row = prefixFinder.findNext().getRow();
  var lineItems = items.getRange(row,4,quantity,11).getValues();
  var itemsRange = calculations.getRange(29,1,quantity,11);
  itemsRange.setValues(lineItems);
  buildHistoricalLineItem();
}

function viewInvoices(){
  var sheet = SpreadsheetApp.getActive();
  var file = DriveApp.getFileById(sheet.getId());
  var folder = file.getParents();
  var folderId = folder.next().getId();
  var modal2 = '<script>window.open("https://drive.google.com/drive/folders/' + folderId + '");google.script.host.close();</script>';
  var interface = HtmlService.createHtmlOutput(modal2)
  SpreadsheetApp.getUi().showModalDialog(interface, 'Opening invoices folder')
}

function buildHistoricalLineItem(){
  var sheet = SpreadsheetApp.getActive();
  var today = new Date();
  var details = sheet.getSheetByName('Details');
  var calculations = sheet.getSheetByName('Calculations');
  var history = sheet.getSheetByName('Historical Data');
  var last = history.getLastRow()+1;
  //Clean up previous Line Items
  var rows = calculations.getRange(5,5,1,1).getValue();
  var sortRange = calculations.getRange(29,1,rows,11);
  sortRange.sort([11]);
  //Get values for Historical Line Item
  var number = details.getRange(5,2,1,1).getValue();
  var legalName = details.getRange(3,2,1,1).getValue();
  var prefix = details.getRange(4,2,1,1).getValue();
  var service = calculations.getRange(29,2,1,1).getValue();
  var currency = calculations.getRange(29,5,1,1).getValue();
  var cycle = calculations.getRange(15,5,1,1).getValue();
  var invoiceNum = calculations.getRange(16,5,1,1).getValue();
  var purCount = calculations.getRange(9,5,1,1).getValue();
  var purAmount = calculations.getRange(10,5,1,1).getValue();
  var retCount = calculations.getRange(11,5,1,1).getValue();
  var retAmount = calculations.getRange(12,5,1,1).getValue();
  var review = calculations.getRange(2,2,1,1).getValue();
  const monthNames = ["January", "February", "March", "April", "May", "June","July", "August", "September", "October", "November", "December"];
  var month = monthNames[today.getMonth()-1]
  //Get ranges for new Historical Line Item
  var setAcctNumber = history.getRange(last,1,1,1);
  var setAcctName = history.getRange(last,2,1,1);
  var setPrefix = history.getRange(last,3,1,1);
  var setItemName = history.getRange(last,4,1,1);
  var setCurrency = history.getRange(last,5,1,1);
  var setCycle = history.getRange(last,6,1,1);
  var setNumber = history.getRange(last,8,1,1);
  var setPurCount = history.getRange(last,9,1,1);
  var setPurAmount = history.getRange(last,10,1,1);
  var setRetCount = history.getRange(last,11,1,1);
  var setRetAmount = history.getRange(last,12,1,1);
  var setReview = history.getRange(last,13,1,1);
  var setMonth = history.getRange(last,14,1,1);
  //Set Historical Line Item
  var test = history.getRange(history.getLastRow(),6,1,1).getValue();
  if (test != cycle){
  setAcctNumber.setValue(number);
  setAcctName.setValue(legalName);
  setPrefix.setValue(prefix);
  setItemName.setValue(service);
  setCurrency.setValue(currency);
  setCycle.setValue(cycle);
  setNumber.setValue(invoiceNum);
  setPurCount.setValue(purCount);
  setPurAmount.setValue(purAmount);
  setRetCount.setValue(retCount);
  setRetAmount.setValue(retAmount);
  setReview.setValue(review);
  setMonth.setValue(month);
    var oldapproval = calculations.getRange(4,2,1,1);
    var oldapprover = calculations.getRange(5,2,1,1);
    var oldtime = calculations.getRange(6,2,1,1);
    var oldfeedback = calculations.getRange(7,2,1,1);
    var newvalue = '';
    oldapproval.setValue(newvalue);
    oldapprover.setValue(newvalue);
    oldtime.setValue(newvalue);
    oldfeedback.setValue(newvalue);
  }
}