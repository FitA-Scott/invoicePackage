function CopyTemplate() {
  var active = SpreadsheetApp.getActive();
  var tab = active.getSheetByName('Details and Calculations');
  var legalname = tab.getRange(5,2,1,1).getValue();
  var clientnumber = tab.getRange(11,2,1,1).getValue();
  var filename = (legalname + "-" + clientnumber + "-Invoicing File");
  var destfolder = DriveApp.getFolderById('1aAQef0Op-BEfjq2F2WKpzn7sf_4hEbUc');
  var newdoc = DriveApp.getFileById(active.getId()).makeCopy(filename, destfolder)
  var newdocid = newdoc.getId();      
        destfolder.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
        return newdocid;        
}

function openNewInvoiceDoc() {
  var folder = DriveApp.getFolderById('1aAQef0Op-BEfjq2F2WKpzn7sf_4hEbUc');
  var newdoc = folder.get
  var docid = CopyTemplate();
  var ui = SpreadsheetApp.getUi();
  var url = "https://docs.google.com/spreadsheets/d/"+docid;
  var html = "<script>window.open('" + url + "');google.script.host.close();</script>";
  var userInterface = HtmlService.createHtmlOutput(html);
      ui.showModalDialog(userInterface, "Opening New Invoice File");
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
      var tab = sheet.getSheetByName('Details and Calculations');
      var newdocid = sheet.getId();
      var companyname = tab.getRange(5,2,1,1).getValue();
      var clientnumber = tab.getRange(11,2,1,1).getValue();
      var taxnumber = tab.getRange(9,2,1,1).getValue();
      var targetsheet = SpreadsheetApp.openById('1WQBEVDTyK8XvTG5BkMJMbqWMyKTf3aYuFjCQPuc23GI');
      var targettab = targetsheet.getSheetByName('Client Master List');
      var targettabdata = targettab.getDataRange();
      var targetcompanyname = targettab.getRange(targettabdata.getLastRow()+1,2,1,1);
      var targetclientnumber = targettab.getRange(targettabdata.getLastRow()+1,1,1,1);
      var targettaxnumber = targettab.getRange(targettabdata.getLastRow()+1,4,1,1);
      var targetdocid = targettab.getRange(targettabdata.getLastRow()+1,6,1,1);
  
      targetcompanyname.setValue(companyname);
      targetclientnumber.setValue(clientnumber);
      targettaxnumber.setValue(taxnumber);
      targetdocid.setValue(newdocid);
    
    }
  else if (button == ui.Button.OK && text != 'masterupdate') {
    ui.alert('Password incorrect.'); }
    else if (button == ui.Button.OK && SpreadsheetApp.getActive().getId() == 'M6KI4FZZCbq3pMjhKpJGSpUVA7j-W5aRr') {
    ui.alert('You cannot add the Template to the Master List.'); }  
  else if (button == ui.Button.CANCEL) {}
  else if (button == ui.Button.CLOSE) {}
  
}