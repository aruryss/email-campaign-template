function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Functions')
    .addItem('Send First Email', 'sendEmail')
    .addItem('Sort Opened Emails', 'sortByOpened')
    .addItem('Send Follow-up Email', 'sendFollowUpEmail')
    .addToUi(); // All the functions for you to run commands from Google Sheet
}

function sendEmail() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("email campaign template"); // YOUR EMAIL CAMPAIGN SHEET NAME
  var range = sheet.getDataRange();
  var data = range.getValues();
  var invalidEmails = [];
  var sender_name = 'YOUR NAME'; // Sender name (it will send from Gmail of Google Account you are using to run the script and save the Apps Script project)
  var subjectline;
  var label = GmailApp.createLabel('email-campaign-2024'); // Create a label to track campaign replies
  
  for (var i = 3; i < data.length; i++) {
    if (!data[i][1]) {
      continue; 
    }
    var receiverName = data[i][0]; // In column A
    var receiverEmail = data[i][1]; // In column B
    var receiverCountry = data[i][2]; // In column C
    subjectline = data[i][3]; // In column D

    if (!isValidEmail(receiverEmail)) {
      invalidEmails.push(receiverEmail); // Store invalid email
      continue; // Skip sending email for this invalid address
    } 

    var template = HtmlService.createTemplateFromFile('bodyOfEmail.html');
    template.name = receiverName; // Sending custom variables to the email body
    template.email = receiverEmail; // template.VARIABLE is the name you used in the HTML file
    template.country = receiverCountry; // receiverVARIABle is the name you declared when collecting data from columns

    var message = template.evaluate().getContent();
    
    GmailApp.sendEmail(receiverEmail, subjectline, "", {
      htmlBody: message,
      name: sender_name
    });
    // Label the email
    var threads = GmailApp.search('subject:"' + subjectline + '" to:' + receiverEmail);
    if (threads.length > 0) {
      threads[0].addLabel(label);
    }

    Utilities.sleep(5000); // To prevent getting flagged by Gmail
  }
  return invalidEmails; // Return list of invalid email addresses
}

function isValidEmail(email) {
  // Regular expression for email validation
  var re = /\S+@\S+\.\S+/;
  return re.test(email);
}

// Handles the get request to the server
function doGet(e) {
  try {
    const method = e.parameter['method'];
    const email = e.parameter['email'];

    if (!method || !email) {
      // Log the issue for debugging
      Logger.log("Method or email parameter is missing.");
      return ContentService.createTextOutput("Missing method or email parameter").setMimeType(ContentService.MimeType.TEXT);
    }

    // Check if the method is 'track'
    if (method === 'track') {
      updateEmailStatus(email);
      return ContentService.createTextOutput("Tracking success").setMimeType(ContentService.MimeType.TEXT);
    } else {
      Logger.log("Invalid method: " + method);
      return ContentService.createTextOutput("Invalid method").setMimeType(ContentService.MimeType.TEXT);
    }
  } catch (error) {
    Logger.log("Error in doGet: " + error.message);
    return ContentService.createTextOutput("Error occurred: " + error.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

function updateEmailStatus(emailToTrack) { // updates the column Email Status
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('email campaign template'); // YOUR EMAIL CAMPAIGN SHEET NAME
  var range = sheet.getDataRange();
  var data = range.getValues();
  
  const headers = data[2];  // Row 3 in the sheet corresponds to index 2 in the array
  const emailOpenedIndex = headers.indexOf('Email Status');  // No need to add 1 because index matches column F (index 5)
  
  if (emailOpenedIndex === -1) {
    Logger.log("Email Status column not found.");
    return;
  }

  for (let i = 3; i < data.length; i++) {  // Start from row 4, index 3
    var email = data[i][1];  // In column B
    
    if (emailToTrack === email) {      
      sheet.getRange(i + 1, emailOpenedIndex + 1).setValue('opened'); // Set 'opened' in the corresponding cell
      break;
    }
  }
}


function sendFollowUpEmail() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("email campaign template"); // YOUR EMAIL CAMPAIGN SHEET NAME
  var range = sheet.getDataRange();
  var data = range.getValues();
  var invalidEmails = [];
  var sender_name = 'YOUR NAME'; 
  var subjectline;

  for (var i = 3; i < data.length; i++) {
    if (data[i][5] != "opened") {
      continue; // Skip if email was not opened
    }
    if (!data[i][1]) {
      continue;  // Skip if no email
    }
    var receiverName = data[i][0];
    var receiverEmail = data[i][1];
    var receiverCountry = data[i][2];
    subjectline = data[i][4]; // Follow-up subject line is in column E

    if (!isValidEmail(receiverEmail)) {
      invalidEmails.push(receiverEmail);
      continue;
    } 

    var template = HtmlService.createTemplateFromFile('bodyOfFollowUpEmail.html');
    template.name = receiverName;
    template.email = receiverEmail;
    template.country = receiverCountry;

    var message = template.evaluate().getContent();
    
    GmailApp.sendEmail(receiverEmail, subjectline, "", {
      htmlBody: message,
      name: sender_name
    });

    Utilities.sleep(5000);
  }
  return invalidEmails;
}

function sortByOpened() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangeToSort = "A2:G1000"; // Adjust based on your data range
  var range = sheet.getRange(rangeToSort);
  range.sort({ column: 5, ascending: true }); // Sort by email status (column 6)
}