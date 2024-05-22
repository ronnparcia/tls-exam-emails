function sendScheduledEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Access the active spreadsheet
  var sheet = ss.getSheets()[0]; // Access the sheet with your data (assuming it's the first sheet, change if needed)
  var data = sheet.getDataRange().getValues(); // Get all the data in the sheet
  
  // Loop through the rows to check if emails should be sent
  // Note:  The loop starts from the third row (index 2) and skips the
  //        first row (NOTICE) and second row (header)
  for (var i = 2; i < data.length; i++) { 
    // Get data from the current row per column. 
    // Note: The column index starts from 0. Second bracket is the column index.
    var status = data[i][0];
    var schedule = data[i][2];
    var recipientEmail = data[i][5];
    var generalExamLink = data[i][7];
    var sectionExamLink = data[i][8];

    // Check if the row has already been processed
    if (status === "Sent" || status === "Failed") {
      Logger.log("Email for " + recipientEmail + " has already been processed.\n\n\n\n");
      continue;
    }

    try {
        Logger.log("Sending email for " + recipientEmail);
        sendEmail(recipientEmail, generalExamLink, sectionExamLink); // Send email
        sheet.getRange(i + 1, 1).setValue("Sent"); // Mark the row as processed
        Logger.log("Email successfully sent for " + recipientEmail + "\n\n\n\n"); // Log success
    } catch (error) {
        // Log the error
        Logger.log("ERROR sending email for " + recipientEmail + "(" + error.message + ")\n\n\n\n");

        // Mark the row as failed
        sheet.getRange(i + 1, 1).setValue("Failed");
    }
  } 
} 

function sendEmail(recipientEmail, generalExamLink, sectionExamLink) {
  // Email subject
  var subject = "[THE LASALLIAN] AY 2023-2024 Term 3 General & Section Exam";

  // Create template from email.html and replace placeholders
  var body = HtmlService.createTemplateFromFile("email");
  body.generalExamLink = generalExamLink;
  body.sectionExamLink = sectionExamLink;

  // Send email
  GmailApp.sendEmail(
    recipientEmail, 
    subject,
    "",
    {
      htmlBody: body.evaluate().getContent(),
      name: "The LaSallian Applications"
    }
  );
}

function isForSending(scheduleDate, currentTime) {
  return (
    scheduleDate.getFullYear() === currentTime.getFullYear() &&
    scheduleDate.getMonth()    === currentTime.getMonth()    &&
    scheduleDate.getDate()     === currentTime.getDate()     &&
    scheduleDate.getHours()    === currentTime.getHours()    &&
    scheduleDate.getMinutes()  === currentTime.getMinutes()
  );
}