function sendScheduledEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Access the active spreadsheet
  var sheet = ss.getSheets()[0]; // Access the sheet with your data (assuming it's the first sheet, change if needed)
  var currentTime = new Date(); // Get the current date and time
  var data = sheet.getDataRange().getValues(); // Get all the data in the sheet
  
  // Loop through the rows to check if emails should be sent
  for (var i = 1; i < data.length; i++) { // Start from the second row assuming headers in the first row
    var schedule = data[i][1];
    var recipientEmail = data[i][4];
    var generalExamLink = data[i][5];
    var sectionExamLink = data[i][6];
     
    var scheduleDate = new Date(schedule); // Convert the schedule cell (assumed to be in 9/27/2023 7:59:00 format) to a Date object

    // Check if the scheduleDate matches the current time
    if (isForSending(scheduleDate, currentTime)) {
      // Attempt to send the email and handle any exceptions
      try {
        // Log attempt to send email
        Logger.log("Sending email for " + recipientEmail);
        Logger.log("Schedule for " + recipientEmail + " is at " + scheduleDate.getFullYear() + "/" + scheduleDate.getMonth() + "/" + scheduleDate.getDate() + " " + scheduleDate.getHours() + ":" + scheduleDate.getMinutes());
        Logger.log("Current time is " + currentTime.getFullYear() + "/" + currentTime.getMonth() + "/" + currentTime.getDate() + " " + currentTime.getHours() + ":" + currentTime.getMinutes());

        // Send email
        sendEmail(recipientEmail, generalExamLink, sectionExamLink);

        // Mark the row as processed (if needed)
        sheet.getRange(i + 1, 1).setValue("✓");
        
        // Log success
        Logger.log("Email successfully sent for " + recipientEmail + "\n\n\n\n");
      } catch (error) {
        // Log the error
        Logger.log("Error sending email for " + recipientEmail + "(" + error.message + ")\n\n\n\n");
      } // Closing brace of catch block
    } // Closing brace of if block
  } // Closing brace of for block
} // Closing brace of function

function sendEmail(recipientEmail, generalExamLink, sectionExamLink) {
  // Email subject
  var subject = "[THE LASALLIAN] AY 2023-2024 Term 1 General & Section Exam";

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