// On install function required for google add on format. 
function onInstall()
{
  onOpen();
}


// Will be used for possible sidebar functionality. 
function onOpen()
{
  DocumentApp.getUi().createAddonMenu().addItem('Finalize Document', 'showSidebar').addToUi();
  
}

//function to display sidebar for add-on
function showSidebar()
{
  var ui = HtmlService.createHtmlOutputFromFile('sidebar').setTitle("Final Doc");
  DocumentApp.getUi().showSidebar(ui);
  
}

function RemoveEditors() {
    //Find active document to be finalized and look up list of editors
  var currentDoc = DocumentApp.getActiveDocument();
  var editors = currentDoc.getEditors();
  var originalName = DocumentApp.getActiveDocument().getName();
  
  //Confirm User is sure they want to removal all editor access. Warn users that all editors will be removed.
   var ui = DocumentApp.getUi();
   var response = ui.alert('Confirm Removal Of Editors', 'Are you sure you want to remove all edit access? \n Doing so will remove all editor access to the doc except for the owner.', ui.ButtonSet.YES_NO);
   Logger.log('User response: ' + response);
   
  //Process the document based on response
 if (response == ui.Button.YES) {
   Logger.log('The user has confirmed removal of editors', response.YES);
   Logger.log(editors);
   
   //Remove Editor Access 
   Logger.log(editors);
   for (var i = 0; i < editors.length; i++) {
     currentDoc.removeEditor(editors[i]);
    
    };
   
  //Send email to the owner with details of the editor removal
   var ownerEmail = Session.getActiveUser().getEmail();
   Logger.log(ownerEmail);
   //var htmlBody = new String("Editors Removed from: " + originalName, "Original File Name" + "\n" + "Original URL " + fileURL);
   MailApp.sendEmail(ownerEmail,"donotreply@edmonton.ca","Editors Removed from " + originalName, "The following editors have been removed from " + originalName + "\n\n" + "Editor List: " + editors + "\n\n" + "Note: The owner of the doc will still be able to edit it.");
   
  //Display Final Message to User telling them about email confirmation. 
  DocumentApp.getUi().alert("Editor Access has been removed","All editor access has been removed.\n\nA confirmation email will be sent to you with additional details.",DocumentApp.getUi().ButtonSet.OK);
  
   
 } else if (response == ui.Button.NO) {
   Logger.log('The user did not confirm removal of editors');
 } else {
   Logger.log('The user clicked the close button in the dialog\'s title bar.');
 }
  
}

//This function will remove editor access, create a final copy of the original doc with no revision history and send an email to the user with the details. 
function FinalizeDoc() {
  
  //Find active document to be finalized and look up list of editors
  var currentDoc = DocumentApp.getActiveDocument();
  var editors = currentDoc.getEditors();
  
  //Confirm User is sure they want to finalize the document. Warn them that revisions and editors will be removed.
   var ui = DocumentApp.getUi();
   var response = ui.alert('Confirm Document Finalization', 'Are you sure you want to finalize this document? \n Doing so will remove all editor access to the original doc and create a new final copy that has no revision history.', ui.ButtonSet.YES_NO);
  
  //Process the document based on response
 if (response == ui.Button.YES) {
   Logger.log('The user has confirmed finalization', response.YES);
   
   //Remove Editior Access 
   Logger.log(editors);
   for (var i = 0; i < editors.length; i++) {
     currentDoc.removeEditor(editors[i]);
    
    };
    //Lookup the info of the original document 
  var fileID = DocumentApp.getActiveDocument().getId();
  var fileURL = DocumentApp.getActiveDocument().getUrl();
  var originalName = DocumentApp.getActiveDocument().getName();
  var newFileName = (DocumentApp.getActiveDocument().getName().concat(" - Final"));
  
  //var filething = DocumentApp.GetActive
  Logger.log(fileID);
  Logger.log(fileURL);
  Logger.log(newFileName);
  
  //Make a copy of the original file
  var fileDriveID = DriveApp.getFileById(fileID);
  var newDriveFile = fileDriveID.makeCopy(newFileName); 
  var newDriveURL = newDriveFile.getUrl();
  Logger.log(newDriveURL);

  //Send email to the owner with details on the orginal and final documents. 
  var ownerEmail = Session.getActiveUser().getEmail();
  var htmlBody = new String("Document Finalized: " + originalName, "Original File Name" + "\n" + "Original URL " + fileURL);
   MailApp.sendEmail(ownerEmail,"donotreply@edmonton.ca","Document Finalized: " + originalName, "Original File Name: " + originalName + "\n" + "Original Document URL: " + fileURL);
   
  //Display Final Message to User telling them about email confirmation. 
  DocumentApp.getUi().alert("Document has been finalized","All editor access has been removed and a final copy of the doc has been successfully created.\n\nA confirmation email will be sent to you with additional details.",DocumentApp.getUi().ButtonSet.OK);
  
   
 } else if (response == ui.Button.NO) {
   Logger.log('The user did not confirm finalization process');
 } else {
   Logger.log('The user clicked the close button in the dialog\'s title bar.');
 }
  

}



