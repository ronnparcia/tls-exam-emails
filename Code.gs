function genSecProcessSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Access the active spreadsheet
  var sheet = ss.getSheetByName(sheetName); // Access the sheet with your data (assuming it's the first sheet, change if needed)
  var data = sheet.getDataRange().getValues(); // Get all the data in the sheet

  // Loop through the rows to check if emails should be sent
  // Note:  The loop starts from the third row (index 2) and skips the
  //        first row (NOTICE) and second row (header)
  for (var i = 2; i < data.length; i++) {
    // Get data from the current row per column.
    // Note: The column index starts from 0. Second bracket is the column index.
    var genSecStatus = data[i][0];
    var recipientEmail = data[i][4];
    var generalExamLink = data[i][6];
    var sectionExamLink = data[i][7];

    // Check if the row has already been processed
    if (genSecStatus === "Sent" || genSecStatus === "Failed") {
      Logger.log(
        "Email for " + recipientEmail + " has already been processed.\n\n\n\n"
      );
      continue;
    }

    try {
      Logger.log("Sending email for " + recipientEmail);
      sendEmail(recipientEmail, generalExamLink, sectionExamLink); // Send email
      sheet.getRange(i + 1, 1).setValue("Sent"); // Mark the row as processed
      Logger.log("Email successfully sent for " + recipientEmail + "\n\n\n\n"); // Log success
    } catch (error) {
      // Log the error
      Logger.log(
        "ERROR sending email for " +
          recipientEmail +
          "(" +
          error.message +
          ")\n\n\n\n"
      );

      // Mark the row as failed
      sheet.getRange(i + 1, 1).setValue("Failed");
    }
  }
}

function genSecMay25Morning() {
  genSecProcessSheet("May 25 Morning");
}

function genSecMay25Afternoon() {
  genSecProcessSheet("May 25 Afternoon");
}

function genSecMay29Morning() {
  genSecProcessSheet("May 29 Morning");
}

function genSecMay29Afternoon() {
  genSecProcessSheet("May 29 Afternoon");
}

function genSecSendEmail(recipientEmail, generalExamLink, sectionExamLink) {
  // Email subject
  var subject = "[THE LASALLIAN] General and Section-Specific Exams";

  // Create template from email.html and replace placeholders
  var body = HtmlService.createTemplateFromFile("genSecEmail");
  body.generalExamLink = generalExamLink;
  body.sectionExamLink = sectionExamLink;

  // Send email
  GmailApp.sendEmail(recipientEmail, subject, "", {
    htmlBody: body.evaluate().getContent(),
    name: "The LaSallian Applications",
  });
}
