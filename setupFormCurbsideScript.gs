/*Copyright 2020 Google LLC
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at
    https://www.apache.org/licenses/LICENSE-2.0
Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.*/

var googleSheetsTemplateID = '1Tw9DUTV1Cr2FTXHgCCUX4ke58Xu5wqo1wN4K0FaqUwg';
var curbsideEmail = 'curbsidepickupsolution@gmail.com';

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

async function copySpreadSheet(e) {
  var date = new Date();
  var formResponse = readForm(e);

  // Make a copy of the template file
  var documentId = await DriveApp.getFileById(googleSheetsTemplateID).makeCopy().getId();

  var ss = SpreadsheetApp.openById(documentId);
  var file = DriveApp.getFileById(documentId);

  var storeName = formResponse.get("What is your Business' Name");
  var storeDescription = formResponse.get("Add a brief description of your business");
  var storePickupAddress = formResponse.get("What is the address of your business that customers will use for the Curbside Pickup?");
  var storeEmail = formResponse.get("What is your business' Email address you want to associate with the Curbside Pickup tool?");
  var storeNumber = formResponse.get("What is the phone number you want to share so that your customers can contact you?");

  // Rename the copied file
  await file.setName(storeName + '\'s Curbside Pickup');

  // Populate the store info sheet
  await populateStoreInfoSheet(ss, storeName, storeDescription, storePickupAddress, storeEmail, storeNumber);

  // Transfer ownership and grant access to the copy of the spreadsheet
  await file.setStarred(true);
  await ss.addViewer(storeEmail);
  await ss.addEditor(storeEmail);

  // Send email to the business with the link for the spreadsheet and instructions on the next steps.
  await sendEmail(storeName, storeEmail, file.getUrl());

  await file.setOwner(storeEmail);
  await file.removeEditor(curbsideEmail);
}

function readForm(e) {
  var responseMap = new Map();
  var itemResponses = e.response.getItemResponses();
  for (var i = 0; i < itemResponses.length; i++) {
    let itemResponse = itemResponses[i];
    responseMap.set(itemResponse.getItem().getTitle(), itemResponse.getResponse());
  }
  return responseMap;
}

async function populateStoreInfoSheet(ss, storeName, storeDescription, storePickupAddress, storeEmail, storeNumber) {
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

async function sendEmail(storeName, storeEmail, spreadsheetUrl){
  MailApp.sendEmail({
      to: storeEmail,
      subject: "Welcome to Curbside Pickup " + storeName + "!",
      htmlBody: "Hi there " + storeName +
      "!<br><br>Great news to have you interested in growing your business during these difficult times by enabling Curbside Pickup!" + 
      "<br><br>We've created a Google Sheet for you where you can have full control over your Curbside Pickup Interface & Orders." +
      "<br>Please click <a href='" + spreadsheetUrl + "'>this link</a> to access your Google Sheet and complete the process (steps below)." +
      "<br><br>Now you just need to complete 4 very simple steps to start getting Curbside Pickup orders:" +
      "<br><br> - 1) Go to the <b>Inventory</b> tab and populate it with your inventory" +
      "<br> - 2) On the <b>Navigation Bar</b> click on <b>Curbside Pickup</b> and then on <b>Authorize Curbside Pickup</b> - Please authenticate with your Google account." +
      "<br>(If you get a warning message indicating that \"This app isn't verified\", click <b>Advanced</b>, <b>Go to OrderSheet Curbside Pickup</b> and <b>Allow</b>)" +
      "<br> - 3) On the <b>Navigation Bar</b> click on <b>Curbside Pickup</b> and then on <b>Update order menu</b> (this will take about a minute to run, and for the first time you run it, you will receive an email with useful information for the last step)" +
      "<br> - 4) Open the email, copy the <b>Customer Order Form</b> link into your social media pages / website and that's it!" +
      "<br><br>You are now Curbside Pickup enabled!" +
      "<br><br><br><i>Disclaimer: This is not an official product, and is made available open-sourced as is under the Apache 2.0 license.</i>"
    });
}