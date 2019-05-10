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
  var startDate = calculationSource.getRange(24,2,1,1).getValue();
  var endDate = calculationSource.getRange(25,2,1,1).getValue();
  var lineItemSheet = SpreadsheetApp.getActive().getSheetByName('Billing Log');
  var invoiceNumberRange = lineItemSheet.getRange(1,2,1,1);
  var accountNameRange = lineItemSheet.getRange(1,3,1,1);
  var lineItemRange = lineItemSheet.getRange(1,4,1,1);
  var startDateRange = lineItemSheet.getRange(1,5,1,1);
  var endDateRange = lineItemSheet.getRange(1,6,1,1);
  var billCurrencyRange = lineItemSheet.getRange(1,8,1,1);
  var nettoAmountRange = lineItemSheet.getRange(1,9,1,1);
  var dueDateRange = lineItemSheet.getRange(1,14,1,1);
  invoiceNumberRange.setValue(invoiceNumber);
  accountNameRange.setValue(accountName);
  lineItemRange.setValue(lineItem);
  startDateRange.setValue(startDate);
  endDateRange.setValue(endDate);
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
  GmailApp.createDraft(deliveryaddresses, emailsubject, emailtext,{ name: 'Fit Analytics GmbH Accounts Receivable', from: 'invoices@fitanalytics.com', replyto: 'invoices@fitanalytics.com', bcc: 'invoices@fitanalytics.com; C.Klawitter@steuerberater-zp.de', attachments:[blob.getAs(MimeType.PDF)],htmlbody:true, footer:true})
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
function mergeTransationData() {
  var sourcesheet = SpreadsheetApp.getActive();
  var sourcetab = sourcesheet.getSheetByName('Purchase Data');
  var sourcerange = sourcetab.getRange(sourcetab.getLastRow(),1,1,12);
  var sourcevalues = sourcerange.getValues();
  var targettab = sourcesheet.getSheetByName('Historical Data');
  var targetMonth = sourcetab.getRange(sourcetab.getLastRow(),6,1,1).getValue();
  var invoiceMonth = sourcesheet.getSheetByName('Invoice').getRange(9,7,1,1);
    
    invoiceMonth.setValue(targetMonth);
    targettab.getRange(targettab.getLastRow()+1,1,1,12).setValues(sourcevalues);
    sourcetab.deleteRow(sourcerange.getRow());  
}