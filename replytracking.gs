function checkForReplies() {
  try {
    var label = GmailApp.getUserLabelByName('email-campaign-2024'); // Get the label used for tracking
    var threads = label.getThreads(); // Get all threads with the label

    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var messages = thread.getMessages();

      for (var j = 0; j < messages.length; j++) {
        var message = messages[j];

        // Check if it's a reply (message is in the inbox and is unread)
        if (message.isInInbox() && message.isUnread() && message.getFrom() !== Session.getActiveUser().getEmail()) {
          var email = message.getFrom();
          updateReplyStatus(email); // Mark reply in the sheet

          // Mark the message as read so it’s not processed again
          message.markRead();
        }
      }
    }
  } catch (error) { // Catch errors of internal lack of resources
    Logger.log('Error in checkForReplies: ' + error.message);
  }
}


function updateReplyStatus(emailToTrack) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("email campaign template");
  var range = sheet.getDataRange();
  var data = range.getValues();

  const headers = data[2]; // Row 3 in the sheet corresponds to index 2 in the array
  const emailRepliedIndex = headers.indexOf('Email Status') + 1;

  for (let i = 3; i < data.length; i++) {  // Start with data in row 4
    var email = data[i][1]; // Column B is index 1
    
    if (emailToTrack.includes(email)) {  // Use includes to match email domain if full address differs slightly
      sheet.getRange(i + 1, emailRepliedIndex).setValue('replied');
      break;
    }
  }
}
