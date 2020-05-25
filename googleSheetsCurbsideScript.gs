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

var ss = SpreadsheetApp.getActive();
var inventorySheet = ss.getSheetByName("Inventory");
var infoSheet = ss.getSheetByName("Info");
var ordersSheet = ss.getSheetByName("Orders");
var formLinksSheet = ss.getSheetByName("Form Links");

var isFirstScriptExecution = false;

function onOpen() {
  var menu = [{name: 'Authorize Curbside Pickup', functionName: 'authorize'}, {name: 'Update Order Menu', functionName: 'main'}];
  SpreadsheetApp.getActive().addMenu('Curbside Pickup', menu);
}

function authorize(){}

function main(){
  //Removing existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  // Creating form for customers to place an order
  var formURL = createOrderForm();
  // Creating form for employees to reply to the order and update the customer
  createOrderUpdateForm();
  
  // If this script is being run for the first time, send welcome email with link to the order form & details on where to find these on the spreadsheet.
  if(isFirstScriptExecution) {
    var infoData = infoSheet.getDataRange().getValues();
    var storeEmail = infoData[1][3];
    var storeName = infoData[1][0];
    sendWelcomeEmail(storeName, storeEmail, ss.getUrl(), formURL);
    isFirstScriptExecution = false;
  }

  // Setting up the scheduler
  ScriptApp.newTrigger("main").timeBased().everyHours(1).create();
}

function createOrderForm() {
  var data = inventorySheet.getDataRange().getValues();
  console.log(data);
  var infoData = infoSheet.getDataRange().getValues();
  console.log(infoData);
  var storeEmail = infoData[1][3];
  var storeName = infoData[1][0];
  // creating the order form
  var orderForm = null;
  var range = formLinksSheet.getRange("B2:B2");
  if (range.isBlank()) { // if this is the first run, create a new form, otherwise reuse the existing one
    orderForm = FormApp.create('Curbside Pickup!');
    // Log the form URL in the spreadsheet
    var range = formLinksSheet.getRange(2, 2, 1, 2);
    var formURLs = [
      [ orderForm.getPublishedUrl(), orderForm.getEditUrl() ]
    ];
    range.setWrap(true); 
    range.setValues(formURLs);
    isFirstScriptExecution = true;
  }
  else {
    var orderFormURL = formLinksSheet.getDataRange().getValues()[1][2]+"";
    orderForm = FormApp.openByUrl(orderFormURL);
  }
  orderForm.setTitle('Curbside Pickup!');

  // Building the form's description
  var openingTime = infoData[1][5];
  var closingTime = infoData[1][6];
  var daysOpen = infoData[1][7];
  var standardPickupTime = infoData[1][8];
  var additionalNotes = infoData[1][9];
  var description = infoData[1][1];
  var menuItemsSentence = description + "\n\nYou can see the below list of all menu items that are currently available." + 
    "\nPlease select how many units of each items you want to order." + 
    "\nPrices don't include HST";

  var detailedDesc = ""
  if (openingTime != "" && closingTime != ""){
    detailedDesc += "We are open from " + openingTime + " to " + closingTime;
  }
  if (daysOpen != ""){
    detailedDesc == "" ? detailedDesc += "We are open " + daysOpen : detailedDesc += ",  " + daysOpen;
  }
  if (standardPickupTime != ""){
    if(detailedDesc != ""){
      detailedDesc += "\n";
    }
    detailedDesc += "Our orders are ready for pickup on an average of " + standardPickupTime + " minutes.";
  }  
  if (additionalNotes != ""){
    detailedDesc == "" ? detailedDesc += additionalNotes : detailedDesc += "\n\n" + additionalNotes;
  }  
  detailedDesc == "" ? detailedDesc += menuItemsSentence : detailedDesc += "\n\n" + menuItemsSentence;

  orderForm.setDescription(detailedDesc);
  // Clear the order form
  var itemsToDelete = orderForm.getItems();
  while (itemsToDelete.length > 0) {
    var itemToDelete = itemsToDelete.pop();
    orderForm.deleteItem(itemToDelete);
  }
  Logger.log('Cleared old version of Order Form');
  orderForm.setTitle(storeName + " Curbside Pickup Menu");
  // Repopulate the order form
  for (i = 1; i < data.length; i++) {
    if ( data[i][2] != "no") {
      let availability = []
      if(data[i][3] >= 5) {
        availability = ['1', '2', '3', '4', '5']
      }
      else {
        for (j = 1; j <= data[i][3]; j++)
          availability.push(j + "")
      }
      console.log(availability);
      let quantity = []
      orderForm.addGridItem()
        .setTitle(data[i][0]+'')
        .setRows(['$' + data[i][1] + ' per item'])
        .setColumns(availability);
    }
  }
  Logger.log('Populated Menu in the Order Form');
  orderForm.addPageBreakItem().setTitle("Personal / Pickup Info:");
  orderForm.addTextItem().setTitle('What is your Name?').setRequired(true);
  orderForm.addTextItem().setTitle('What is your email address?').setRequired(true);
  orderForm.addTextItem().setTitle('What is your phone number?').setRequired(true);
  orderForm.addTextItem().setTitle('Do you have any specific comments?');
  Logger.log('Populated Personal Details in the Order Form');
  // Creating the trigger for sending emails on order form submissions
  var x = ScriptApp.newTrigger('onOrderFormSubmit').forForm(orderForm).onFormSubmit().create();
  Logger.log("Created Order Form");
  return orderForm.getPublishedUrl();
}

function onOrderFormSubmit(e) {
  var infoData = infoSheet.getDataRange().getValues();
  var storeEmail = infoData[1][3];
  var storeName = infoData[1][0];
  var storeNumber = infoData[1][4];
  var storeAddress = infoData[1][2];
  let responseMap = new Map();
  let responseOrdersSet = new Set();
  var stillInTheMenu = true;
  var itemResponses = e.response.getItemResponses();
  for (var i = 0; i < itemResponses.length; i++) {
    var itemResponse = itemResponses[i];
    responseMap.set(itemResponse.getItem().getTitle(), itemResponse.getResponse())
    if(itemResponse.getItem().getTitle() === "What is your Name?"){
      stillInTheMenu = false;
    }
    if(stillInTheMenu){
      responseOrdersSet.add([itemResponse.getItem().getTitle(), itemResponse.getResponse()]);
    }
  }
  var name = responseMap.get("What is your Name?")
  var email = responseMap.get("What is your email address?")
  var phoneNumber = responseMap.get("What is your phone number?")
  var comments = responseMap.get("Do you have any specific comments?")
  
  // Accounting for the scennario where the customer didn't select anything in the menu - ask the customer to submit another request.
  if(responseOrdersSet.size == 0){
    var orderFormPublicURL = formLinksSheet.getDataRange().getValues()[1][1];
    // Send Customer Email
    MailApp.sendEmail({
      to: email,
      subject: "Sorry! There was an issue with your order " + name + " at " + storeName,
      htmlBody: "Hi " + name + "!<br><br>We've got your order request but unfortunately we did not record any item selection from our menu." + 
      "<br><br>Please click <a href='" + orderFormPublicURL + "'>this link</a> to submit a new order." +
      "<br><br>---------------------<br><br>" + storeName + "<br>" + storeNumber + "<br>" + storeEmail + "<br>" + storeAddress + "<br>"
    });
  }
  else{
    var orderItemsStringEmailFormat = "<br>";
    var orderItemsStringSheetsFormat = "=CONCATENATE(";
    responseOrdersSet.forEach(item => {
      orderItemsStringEmailFormat += "<b>   x " + item[1] + "</b>  " + item[0] + "<br>";
      orderItemsStringSheetsFormat += "\" x "+ item[1] + " " + item[0] + "\", CHAR(10), ";
    });
    orderItemsStringSheetsFormat = orderItemsStringSheetsFormat.substr(0, orderItemsStringSheetsFormat.length-12);
    orderItemsStringSheetsFormat += ")";
    // Log the order in the spreadsheet
    var range = ordersSheet.getRange(ordersSheet.getLastRow() + 1, 1, 1, 7);
    var d = new Date();
    var timeStamp = d.getTime(); 
    var currentTime = d.toLocaleString();
    var order = [
      [ timeStamp, currentTime, orderItemsStringSheetsFormat, comments, name, email, phoneNumber ]
    ];
    range.setWrap(true); 
    range.setValues(order);
    var updateFormPublicURL = formLinksSheet.getDataRange().getValues()[2][1];
    // Send Customer Email
    MailApp.sendEmail({
      to: email,
      subject: "Thank you for your order " + name + "!",
      htmlBody: "Hi " + name + "!<br><br>We've got your order and are working through it as fast as we can.<br> Your order ID is <b>" + timeStamp + "</b>.<br>We will send you an update once you know the exact pick up time.<br><br>Here are the contents of your order:" + orderItemsStringEmailFormat + "<br><br>---------------------<br><br>Please don't hesitate to reach out to use in case you have any questions, and Thank you for your order!!<br><br>" + storeName + "<br>" + storeNumber + "<br>" + storeEmail + "<br>" + storeAddress + "<br>"
    });
    // Send Store Email
    MailApp.sendEmail({
      to: storeEmail,
      subject: "New order for " + storeName + "!",
      htmlBody: "Hi " + storeName + "!<br><br>You've got a new order (order ID: <b>" + timeStamp + "</b>): " + orderItemsStringEmailFormat + "<br><br>Please click <a href='" + updateFormPublicURL + "'>this link</a> to update the customer on the pickup time."
    });
  }
  Logger.log("Order submitted & Emails sent.");
}

function createOrderUpdateForm() {    
  var infoData = infoSheet.getDataRange().getValues();
  var storeName = infoData[1][0];
  // creating the order form
  var updateForm = null;
  var range = formLinksSheet.getRange("B3:B3");
  if (range.isBlank()) { // if this is the first run, create a new form, otherwise reuse the existing one
    updateForm = FormApp.create('Curbside Pickup! Order Update');
    // Log the form URL in the spreadsheet
    var range = formLinksSheet.getRange(3, 2, 1, 2);
    var formURLs = [
      [ updateForm.getPublishedUrl(), updateForm.getEditUrl() ]
    ];
    range.setWrap(true); 
    range.setValues(formURLs);
  }
  else {
    var updateFormURL = formLinksSheet.getDataRange().getValues()[2][2]+"";
    updateForm = FormApp.openByUrl(updateFormURL);
  }
  updateForm.setTitle('Curbside Pickup! Order Update');
  updateForm.setDescription("Please fill in the fields below to update your customer on their order");
  // Clear the updateForm
  var itemsToDelete = updateForm.getItems();
  while (itemsToDelete.length > 0) {
    var itemToDelete = itemsToDelete.pop();
    updateForm.deleteItem(itemToDelete);
  }
  Logger.log('Cleared old version of Order Update Form');
  updateForm.setTitle(storeName + " Menu");
  updateForm.addTextItem().setTitle('What is the order ID?').setRequired(true);
  updateForm.addTextItem().setTitle('What is the updated pickup time?').setRequired(true);
  updateForm.addTextItem().setTitle('Add any specific messages here');
  Logger.log('Populated updateForm Details');
  // Log the form URL in the spreadsheet
  var range = formLinksSheet.getRange(3, 2, 1, 2);
  var formURLs = [
    [ updateForm.getPublishedUrl(), updateForm.getEditUrl() ]
  ];
  range.setWrap(true); 
  range.setValues(formURLs);
  // Creating the trigger for sending emails on form submissions
  ScriptApp.newTrigger('onOrderUpdateFormSubmit').forForm(updateForm).onFormSubmit().create();
  Logger.log("Created Order Update Form");
}

function onOrderUpdateFormSubmit(e) {
  let responseMap = new Map();
  var itemResponses = e.response.getItemResponses();
  for (var i = 0; i < itemResponses.length; i++) {
    var itemResponse = itemResponses[i];
    responseMap.set(itemResponse.getItem().getTitle(), itemResponse.getResponse())
  }
  var orderID = responseMap.get("What is the order ID?")
  var updatedPickuptime = responseMap.get("What is the updated pickup time?")
  var restaurantComments = responseMap.get("Add any specific messages here")
  // finding the correspondant email address
  var orders = ordersSheet.getDataRange().getValues();
  var emailAddress;
  var customerName;
  for (i = 1; i < orders.length; i++) {    
    if (orders[i][0]+"" == orderID+"") {
      emailAddress = orders[i][5]+"";
      customerName = orders[i][4]+"";
    }
  }   
  var infoData = infoSheet.getDataRange().getValues();
  var storeEmail = infoData[1][3];
  var storeName = infoData[1][0];
  var storeNumber = infoData[1][4];
  var storeAddress = infoData[1][2];
  var emailBody = "Hi there " + customerName + "!<br><br>We wanted to let you know that we've started preparing your order with ID <b>" + orderID +"</b> and it will be ready for pickup by <b>" + updatedPickuptime + "</b>!";
  if(restaurantComments != ""){
    emailBody += "<br><br>Note: " + restaurantComments;
  }
  emailBody += "<br><br>Please don't hesitate to reach out to use in case you have any questions, and Thank you for your order!!<br><br>" + storeName + "<br>" + storeNumber + "<br>" + storeEmail + "<br>" + storeAddress + "<br>";
  MailApp.sendEmail({
    to: emailAddress,
    subject: "Thank you for your order!",
    htmlBody: emailBody
  });
  Logger.log("Order Updated Email sent.");
}

function sendWelcomeEmail(storeName, storeEmail, spreadsheetUrl, formURL) {
  MailApp.sendEmail({
      to: storeEmail,
      subject: "Welcome to Curbside Pickup " + storeName + " - Your Order Form is ready to be shared with your customers!",
      htmlBody: "Hi there " + storeName + "!" +
      "<br><br>We have just completed the creation of the Curbside Pickup Order form." +
      "<br>To access it please click <a href='" + formURL + "'>this link</a>" +
      "<br>And this is its link which you can share on your social media / website: " + formURL +
      "<br><br>Don't forget about <a href='" + spreadsheetUrl + "'>your Google Sheet</a> where you can keep control over your Inventory & Orders." +
      "<br><br>If you navigate to the tab <b>Form Links</b> you will find four links:" +
      "<br>- 1) <b>Order Submission Form - Published Form URL</b>: This is the link you can share with your customers so that they can submit orders." +
      "<br>- 2) <b>Order Submission Form - Form Edit URL</b>: This is the link you can use to customize your Customer Order Submission form (you can tailor the color scheme to your business or add a relevant background image)." +
      "<br>- 3) <b>Order Update Form - Published Form URL</b>: This is the link you can use to access the form to update your customers on their orders." +
      "<br>- 4) <b>Order Update Form - Form Edit URL</b>: This is the link you can use to customize your Order Update Form." +
      "<br><br>Note on the Form Edit URLs: You should only share the \"Order Submission Form - Published Form URL\"link with your customers and the two Form Edit URL links are intended for changes to the color scheme / background images only. You should not use their content as it is controlled through the Google Sheet and your manual changes will be overwritten." +
      "<br><br><b>Remember, you agree that you are responsible for the information you collect about your customers (including their names and contact information). Please keep this information secure.</b>" +
      "<br><br>Welcome, and we really hope you enjoy this functionality!" +
      "<br><br><br><i>Disclaimer: This is not an official product, and is made available open-sourced as is under the Apache 2.0 license.</i>"
    });
}