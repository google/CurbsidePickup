var googleSheetsTemplateID = '1jIBa7AdzrMOa13iZF6A7M9W2WZr_8YBRLa1SbZtiulM'

var inventorySheet = "Inventory";
var infoSheet = "Info";
var ordersSheet = "Orders";
var formLinksSheet = "Form Links";

function main() {
  // Removing existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  // Creating the trigger to copy the spreadsheet
  var x = ScriptApp.newTrigger('copySpreadSheet').forForm(FormApp.getActiveForm()).onFormSubmit().create();
}

function copySpreadSheet(e) {
  var date = new Date();
  let formResponse = readForm(e);  

  // Make a copy of the template file
  var documentId = DriveApp.getFileById(googleSheetsTemplateID).makeCopy().getId();
 
  var ss = SpreadsheetApp.openById(documentId);
  var file = DriveApp.getFileById(documentId);
  
  var storeName = formResponse.get("What is your Business' Name");
  var storeDescription = formResponse.get("Add a brief description of your business");
  var storePickupAddress = formResponse.get("What is the address of your business that customers will use for the Curbside Pickup?");
  var storeEmail = formResponse.get("What is your business' Email address you want to associate with the Curbside Pickup tool?");
  var storeNumber = formResponse.get("What is the phone number you want to share so that your customers can contact you?");

  // Rename the copied file
  file.setName(storeName + '\'s Curbside Pickup');
  Logger.log('document Id: ' + documentId);
  Logger.log('Download Url: ' + file.getDownloadUrl());
  Logger.log('View Url: ' + file.getUrl());

  // Populate the store info sheet
  populateStoreInfoSheet(ss, storeName, storeDescription, storePickupAddress, storeEmail, storeNumber);


  // TODO: bypass the notification email
  // Drive.Permissions.insert(
  //  {
  //    'role': 'owner',
  //    'type': 'user',
  //    'value': storeEmail
  //  },
  //  documentId,
  //  {
  //    'sendNotificationEmails': 'false'
  //  });

  // Transfer ownership and grant access to the copy of the spreadsheet
  file.setStarred(true);
  ss.addViewer(storeEmail);
  ss.addEditor(storeEmail);

  // Send email to the business with the link for the spreadsheet and instructions on the next steps.
  sendEmail(storeName, storeEmail, file.getUrl()); 

  // TODO - uncomment once github process is completed.
  //file.setOwner(storeEmail);
}

function readForm(e) {
  let responseMap = new Map();
  var itemResponses = e.response.getItemResponses();
  for (var i = 0; i < itemResponses.length; i++) {
    var itemResponse = itemResponses[i];
    responseMap.set(itemResponse.getItem().getTitle(), itemResponse.getResponse());
  }
  return responseMap;
}

function populateStoreInfoSheet(ss, storeName, storeDescription, storePickupAddress, storeEmail, storeNumber) {
  // Log the store info in the spreadsheet
  var range = getSpecificSheet(ss,infoSheet).getRange(2, 1, 1, 5);
  var storeInfo = [
    [ storeName, storeDescription, storePickupAddress, storeEmail, storeNumber]
  ];
  range.setWrap(true); 
  range.setValues(storeInfo);

}

function getSpecificSheet(ss, sheetName){
  return ss.getSheetByName(sheetName);
}

function sendEmail(storeName, storeEmail, spreadsheetUrl){
  MailApp.sendEmail({
      to: storeEmail,
      subject: "Welcome to Curbside Pickup " + storeName + "!",
      htmlBody: "Hi there " + storeName + 
      "!<br><br>Great news to have you interested in growing your business during these dificult times by enabling Curbside Pickup!" + 
      "<br><br>We've created a Google Sheet for you where you can have full control over your Curbside Pickup Interface & Orders." +
      "<br>Please click <a href='" + spreadsheetUrl + "'>this link</a> to access your Google Sheet and complete the process (steps below)." +
      "<br><br>Now you just need to complete 4 very simple steps to start getting Curbside Pickyp orders:" +
      "<br><br> - 1) Go to the <b>Inventory</b> tab and populate it with your inventory" +
      "<br> - 2) On the <b>Navigation Bar</b> click on <b>Curbside Pickup</b> and then on <b>Authorize Curbside Pickup</b> - Please authenticate with your Google account." +
      "<br>(If you get a warning message indicating that \"This app isn't verified\", click <b>Advanced</b>, <b>Go to OrderSheet Curbside Pickup</b> and <b>Allow</b>)" +
      "<br> - 3) On the <b>Navigation Bar</b> click on <b>Curbside Pickup</b> and then on <b>Update order menu</b> (this will take about a minute to run, and for the first time you run it, you will receive an email with useful information for the last step)" +
      "<br> - 4) Open the email, copy the <b>Customer Order Form</b> link into your social media pages / website and that's it!" +
      "<br><br>You are now Curbside Pickup enabled!"
    });
}

// TODO edit the main script to send an email with all the form links in the first time the script is run


















