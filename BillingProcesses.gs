// Global Variables
  var ui = SpreadsheetApp.getUi();
  var active = SpreadsheetApp.getActiveSpreadsheet();
  var detail = active.getSheetByName('Invoice');

function onOpen(e) {
    ui.createMenu('Billing')
      .addItem('Create PDF', 'savePDF')
      .addItem('Create Historical Invoice', 'userInput')
      .addToUi();
      mergeTransactionData();
      showSidebar();
       searchNumber();     
  var testRange = active.getSheetByName('Details and Calculations').getRange(38,1,1,1).getValue();
    if ( testRange == 'Special Item') {
      specialSalesData();
  }
} 
function showSidebar() {
  var list = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Invoice Functions')
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(list);
}

function userInput() {
  var modal = HtmlService.createHtmlOutputFromFile('Details')
      .setWidth(450)
      .setHeight(300);    
  var dialog = ui.showModalDialog(modal, 'Enter Invoicing Details');
}

function logData(billmonth, billyear, adjustment) {
    // Logger.log('Concatenated data is: ' + data);
    //  SpreadsheetApp.getActiveSheet();
    var sheet = SpreadsheetApp.openById("1OJHsuX-7H1qrQdBntPzxwA97QlgBr9jCWOnDlw4VJIo");
    //Select the tab you want
    SpreadsheetApp.setActiveSheet(sheet.getSheets()[1]);
    //Select reference range in the tab (you might need the concept of "last-current-cell"/"current-cell")
    var targetCellMonth = sheet.getRange("B28");
    var targetCellYear = sheet.getRange("B27");
    var targetCellAdj = sheet.getRange("B30");
    //Overwrite the value
    targetCellMonth.setValue(billmonth);
    targetCellYear.setValue(billyear);
    targetCellAdj.setValue(adjustment);
}
function savePDF( optSSId, optSheetId ) {
  var calculationSource = SpreadsheetApp.getActive().getSheetByName('Details and Calculations');
  var invoiceSource = SpreadsheetApp.getActive().getSheetByName('Invoice');
  var invoiceNumber = invoiceSource.getRange(8,7,1,1).getValue();
  var accountNameText = invoiceSource.getRange(27,1,1,1).getValue();
  var accountName = accountNameText.replace(" ","_",'g');
  var dueDate = calculationSource.getRange(47,2,1,1).getValue();
  var nettoAmount = invoiceSource.getRange(31,7,1,1).getValue();
  var billCurrency = calculationSource.getRange(45,2,1,1).getValue();
  var servicePeriodText = invoiceSource.getRange(9,7,1,1).getValue();
  var servicePeriod = servicePeriodText.replace(" ","_");
  var description = invoiceSource.getRange(27,2,1,1).getValue();  
  var product = invoiceSource.getRange(27,3,1,1).getValue();
  var lineItem = description + ' - ' + product + ' // ' + servicePeriod
  var startDate = calculationSource.getRange(22,2,1,1).getValue();
  var endDate = calculationSource.getRange(23,2,1,1).getValue();
  var taxRate = calculationSource.getRange(14,2,1,1).getValue();
  var lineItemSheet = SpreadsheetApp.getActive().getSheetByName('Billing Log');
  var invoiceNumberRange = lineItemSheet.getRange(1,2,1,1);
  var accountNameRange = lineItemSheet.getRange(1,3,1,1);
  var lineItemRange = lineItemSheet.getRange(1,4,1,1);
  var startDateRange = lineItemSheet.getRange(1,5,1,1);
  var endDateRange = lineItemSheet.getRange(1,6,1,1);
  var taxRateRange = lineItemSheet.getRange(1,7,1,1);
  var billCurrencyRange = lineItemSheet.getRange(1,8,1,1);
  var nettoAmountRange = lineItemSheet.getRange(1,9,1,1);
  var dueDateRange = lineItemSheet.getRange(1,14,1,1);
  invoiceNumberRange.setValue(invoiceNumber);
  accountNameRange.setValue(accountName);
  lineItemRange.setValue(lineItem);
  startDateRange.setValue(startDate);
  endDateRange.setValue(endDate);
  taxRateRange.setValue(taxRate);
  nettoAmountRange.setValue(nettoAmount);
  dueDateRange.setValue(dueDate);
  billCurrencyRange.setValue(billCurrency);

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
    var blob = response.getBlob().setName('FitAnalytics_' + invoiceNumber + '_' + accountName + '_' + servicePeriod + '.pdf');
    folder.createFile(blob);
  }
    var result = ui.alert(
      'Would you like to e-mail the invoice?',
      ui.ButtonSet.YES_NO_CANCEL);
  // Process user response.
  if (result == ui.Button.YES) {
    ui.alert('An e-mail draft has been created please proofread and send.');
  //Prepare draft to send to client and second email to send to Digi-Bel and Salesforce
  var infosheet = SpreadsheetApp.getActive();
  var infosource = infosheet.getSheetByName('Invoice');
  var detailsource = infosheet.getSheetByName('Details and Calculations');
  var companyname = infosource.getRange(7,1,1,1).getValues();
  var invoiceperiod = infosource.getRange(9,7,1,1).getValues();
  // Include cc address in GmailApp options when cc address is used for the invoice.
  var ccaddress = detailsource.getRange(52,2,1,1).getValues();
  var bccaddress = detailsource.getRange(53,2,1,1).getValues();
  var deliveryaddresses = detailsource.getRange(4,2,1,1).getValues();
  var emailsubject = detailsource.getRange(49,1,1,1).getValues();
  var emailtext = detailsource.getRange(50,1,1,1).getValues();
  var sheetbodytext = detailsource.getRange(50,1,1,1).getValues(); //returns a two-dimensional array, the text is in the first item
  var emailtext = String(sheetbodytext[0]);
  var emailfooter = ('<div><br><br><br>Mit freundlichen Grüßen / Best regards</div><br><b>Fit Analytics Accounting Team</b><br><br><img src="https://ci5.googleusercontent.com/proxy/92ywHWBtnnjrrcbYhVDoqWjHZNDKD2ukCvaIDfIoFxERJKyIfwLaSW13NVs2ECuVzo63kHv6ZIpZMuPWjBlr28gADggLhp-h4p5qhcQ37au1-aDY2xQTaB9sOGNKtkGk3Rvs5Ze8Xv4C4rjPmYfSrp__0mwmpG5q0THAh84N8eiA3K1HnYXb4OnvuZC4IOZKlJXTDZs64C8=s0-d-e1-ft#https://docs.google.com/uc?export=download&amp;id=0B0gpnzRVY698NUN3WGJoWEk1NXc&amp;revid=0B0gpnzRVY698aUFoUitYeDNpQTRCNWtqTW9VWEtkbGlmK2lJPQ" alt="" width="164" height="32" style="font-family:arial,helvetica,sans-serif;font-size:12.8px" class="CToWUd"></div><br><div>Voigtstraße 3 | 10247 Berlin</div><br><div>www.fitanalytics.com</div>');
  GmailApp.createDraft(deliveryaddresses, emailsubject,'',{ name: 'Fit Analytics GmbH Accounts Receivable', from: 'invoices@fitanalytics.com', replyto: 'invoices@fitanalytics.com', htmlBody: emailtext + emailfooter, bcc: 'invoices@fitanalytics.com; C.Klawitter@steuerberater-zp.de', attachments:[blob.getAs(MimeType.PDF)]});  
  //MailApp.sendEmail('kyle@fitanalytics.com','Our invoice for ' + invoiceperiod, 'Invoice PDF Attached.', { name: 'Fit Analytics GmbH Accounts Receivable', attachments:[blob.getAs(MimeType.PDF)]}); 
  moveBillingLogLineItem()
  // Process alternate user response  
  } else if (result == ui.Button.NO) {
    ui.alert('Print invoice before closing');
    moveBillingLogLineItem()
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
  var destinationRange = destinationSheet.getRange(destinationSheet.getLastRow()+1,1,1,14);
  var billingLogLineItem = SpreadsheetApp.getActive().getSheetByName('Billing Log').getRange(1,1,1,14).getValues();
      destinationRange.setValues(billingLogLineItem);
}
function mergeTransactionData() {
  var sourcesheet = SpreadsheetApp.getActive();
  var sourcetab = sourcesheet.getSheetByName('Purchase Data');
  var sourcerange = sourcetab.getRange(sourcetab.getLastRow(),1,1,12);
  var testCell = sourcetab.getRange(sourcetab.getLastRow(),1,1,1).getValue();
  var sourcevalues = sourcerange.getValues();
  var targettab = sourcesheet.getSheetByName('Historical Data');
  var targetMonth = sourcetab.getRange(sourcetab.getLastRow(),6,1,1).getValue();
  var invoiceMonth = sourcesheet.getSheetByName('Invoice').getRange(9,7,1,1);
    if ( testCell != "Account Number"){
    invoiceMonth.setValue(targetMonth);
    targettab.getRange(targettab.getLastRow()+1,1,1,12).setValues(sourcevalues);
    sourcetab.deleteRow(sourcerange.getRow());
  }
}
function importCustomerData() {
  var activeSheet = SpreadsheetApp.getActive().getSheetByName('Details and Calculations')
  var detailSheet = SpreadsheetApp.openById('1WQBEVDTyK8XvTG5BkMJMbqWMyKTf3aYuFjCQPuc23GI').getSheetByName('Client Info Update');  
  var updateRowNum = searchNumber(); 
  var updateInfo = detailSheet.getRange(updateRowNum,1,1,26).getValues();
  var updateSheet = SpreadsheetApp.getActive().getSheetByName('Customer Data');
  var updateDest = updateSheet.getRange(updateSheet.getLastRow()+1,1,1,26);
  updateDest.setValues(updateInfo);
  refreshCustomerData();
}

function searchNumber() {
  var originSheet = SpreadsheetApp.getActive().getSheetByName('Details and Calculations');
  var temp = originSheet.getRange(53,1,1,1);
  var valuesSheet = SpreadsheetApp.openById('1WQBEVDTyK8XvTG5BkMJMbqWMyKTf3aYuFjCQPuc23GI').getSheetByName('Client Info Update');  
  var accountNumber = originSheet.getRange(11,2,1,1).getValue();
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
  var calcsheet = SpreadsheetApp.getActive().getSheetByName('Details and Calculations');
  var currentmonth = calcsheet.getRange(21,2,1,1).getValue();
  var purchaseFormula = '=IMPORTRANGE("https://docs.google.com/spreadsheets/d/14j0ZpxIBU85hUtRwCrhlQ73N30MsXTMu3LACG7RY4iw/edit", ' + '"Purchases ' + currentmonth + '!A2:G")';
  var returnFormula = '=IMPORTRANGE("https://docs.google.com/spreadsheets/d/14j0ZpxIBU85hUtRwCrhlQ73N30MsXTMu3LACG7RY4iw/edit", ' + '"Returns ' + currentmonth + '!A2:G")';
  var purchaseFormulaDest = itemsheet.getRange(3,1,1,1);
  var returnFormulaDest = itemsheet.getRange(3,9,1,1);
    purchaseFormulaDest.setValue(purchaseFormula);
    returnFormulaDest.setValue(returnFormula);  
}

function refreshCustomerData() {
  var current = SpreadsheetApp.getActive();
  var inbound = current.getSheetByName('Customer Data');
  var destination = current.getSheetByName('Details and Calculations');
  var refreshDate = SpreadsheetApp.getActive().getSheetByName('Invoice').getRange(7,7,1,1).getValue();
  var newRefreshDate = inbound.getRange(inbound.getLastRow(),20,1,1);
  var userName = Session.getEffectiveUser();
  var newUserName = inbound.getRange(inbound.getLastRow(),21,1,1);
  //New values imported form Salesforce report
  var billingEmails = inbound.getRange(inbound.getLastRow(),10,1,1).getValue();
  var billingContacts = inbound.getRange(inbound.getLastRow(),9,1,1).getValue();  
  var legalName = inbound.getRange(inbound.getLastRow(),5,1,1).getValue();
  var addressOne = inbound.getRange(inbound.getLastRow(),6,1,1).getValue();
  var addressTwo = inbound.getRange(inbound.getLastRow(),7,1,1).getValue();
  var addressThree = inbound.getRange(inbound.getLastRow(),8,1,1).getValue();
  var vatId = inbound.getRange(inbound.getLastRow(),4,1,1).getValue();
  var poNumber = inbound.getRange(inbound.getLastRow(),14,1,1).getValue();
  var cpoRate = inbound.getRange(inbound.getLastRow(),16,1,1).getValue();
  var fixedFee = inbound.getRange(inbound.getLastRow(),15,1,1).getValue();
  var paymentTerms = inbound.getRange(inbound.getLastRow(),13,1,1).getValue();
  var currency = inbound.getRange(inbound.getLastRow(),12,1,1).getValue();
  var billMethod = inbound.getRange(inbound.getLastRow(),11,1,1).getValue();
  var cpoMin = inbound.getRange(inbound.getLastRow(),17,1,1).getValue();
  var cpoMax = inbound.getRange(inbound.getLastRow(),18,1,1).getValue();
  // Destinations for the new values
  var newBillingEmails = destination.getRange(4,2,1,1);
  var newBillingContacts = destination.getRange(3,2,1,1);
  var newLegalName = destination.getRange(5,2,1,1);
  var newAddressOne = destination.getRange(6,2,1,1);
  var newAddressTwo = destination.getRange(7,2,1,1);
  var newAddressThree = destination.getRange(8,2,1,1);
  var newVatId = destination.getRange(9,1,1,1);
  var newPoNumber = destination.getRange(10,2,1,1);
  var newCpoRate = destination.getRange(12,2,1,1);
  var newFixedFee = destination.getRange(13,2,1,1);
  var newPaymentTerms = destination.getRange(15,2,1,1);
  var newCurrency = destination.getRange(45,2,1,1);
  var newBillMethod = destination.getRange(46,2,1,1);
  var newCpoMin = destination.getRange(23,2,1,1);
  var newCpoMax = destination.getRange(22,2,1,1);
  // Constant Values that should be moved every time the refresh is run
  newBillingEmails.setValue(billingEmails);
  newBillingContacts.setValue(billingContacts);
  newLegalName.setValue(legalName);
  newAddressOne.setValue(addressOne);
  newAddressTwo.setValue(addressTwo);
  newAddressThree.setValue(addressThree);
  newPaymentTerms.setValue(paymentTerms);
  newCurrency.setValue(currency);
  newBillMethod.setValue(billMethod);
  newRefreshDate.setValue(refreshDate);
  newUserName.setValue(userName);
  // Variable values that vary based on billing method and company invoice requirements
  if (vatId != null) {
    newVatId.setValue(vatId);
    }
  if (poNumber != null) {
    newPoNumber.setValue(poNumber);
    }
  if (cpoRate != null) {
    newCpoRate.setValue(cpoRate);
    }  
  if (fixedFee != null) {
    newFixedFee.setValue(fixedFee);
    }
  if (cpoMin != null) {
    newCpoMin.setValue(cpoMin);
    }
  if (cpoMax != null) {
    newCpoMax.setValue(cpoMax);
    }
}