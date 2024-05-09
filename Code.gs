function sendScheduledEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Access the active spreadsheet
  var sheet = ss.getSheets()[0]; // Access the sheet with your data (assuming it's the first sheet, change if needed)
  var currentTime = new Date(); // Get the current date and time
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
     
    // Convert the schedule cell (assumed to be in 9/27/2023 7:59:00 format) to a Date object
    var scheduleDate = new Date(schedule); 

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

        // Mark the row as processed
        sheet.getRange(i + 1, 1).setValue("Sent");
        
        // Log success
        Logger.log("Email successfully sent for " + recipientEmail + "\n\n\n\n");
      } catch (error) {
        // Log the error
        Logger.log("ERROR sending email for " + recipientEmail + "(" + error.message + ")\n\n\n\n");

        // Mark the row as failed
        sheet.getRange(i + 1, 1).setValue("Failed");
      } 
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