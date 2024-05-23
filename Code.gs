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
      genSecSendEmail(recipientEmail, generalExamLink, sectionExamLink); // Send email
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

// Take Home Exams

function takeHomeProcessSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Access the active spreadsheet
  var sheet = ss.getSheetByName(sheetName); // Access the sheet with your data (assuming it's the first sheet, change if needed)
  var data = sheet.getDataRange().getValues(); // Get all the data in the sheet

  // Loop through the rows to check if emails should be sent
  // Note:  The loop starts from the third row (index 2) and skips the
  //        first row (NOTICE) and second row (header)
  for (var i = 2; i < data.length; i++) {
    // Get data from the current row per column.
    // Note: The column index starts from 0. Second bracket is the column index.
    var recipientEmail = data[i][4];
    var sectionName = data[i][3];
    var takeHomeExamLink = takeHomeGetLink(sectionName);

    try {
      Logger.log("Sending email for " + recipientEmail +"\n");
      Logger.log("Section: " + sectionName);
      Logger.log("Link: " + takeHomeExamLink);
      takeHomeSendEmail(recipientEmail, takeHomeExamLink)
      sheet.getRange(i + 1, 2).setValue("Sent"); // Mark the row as processed
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
      sheet.getRange(i + 1, 2).setValue("Failed");
    }
  }
}

function takeHomeGetLink(sectionName) {
  // Switch case for section name
  var takeHomeExamLink;

  switch (sectionName) {
    case "University":
      takeHomeExamLink = "univ.com";
      break;
    case "Menagerie":
      takeHomeExamLink = "menage.com";
      break;
    case "Sports":
      takeHomeExamLink = "sports.com";
      break;
    case "Vanguard":
      takeHomeExamLink = "vangie.com";
      break;
    case "Intermedia":
      takeHomeExamLink = "intermedia.com";
      break;
    case "Photo":
      takeHomeExamLink = "photo.com";
      break;
    case "Art & Graphics":
      takeHomeExamLink = "a&g.com";
      break;
    case "Layout":
      takeHomeExamLink = "layout.com";
      break;
    case "Web (Web)":
      takeHomeExamLink = "web.com";
      break;
    case "Web (WebDev)":
      takeHomeExamLink = "webdev.com";
      break;
    default:
      takeHomeExamLink = "Error. Please contact us.";
      break;
  }

  return takeHomeExamLink;
}

function takeHomeSendEmail(recipientEmail, takeHomeExamLink) {
    // Email subject
    var subject = "[THE LASALLIAN] AY 2023-2024 Term 3  Take Home Exam";
  
    // Create template from takeHomeExamEmail.html and replace placeholders
    var body = HtmlService.createTemplateFromFile("takeHomeEmail");
    body.takeHomeExamLink = takeHomeExamLink;
  
    // Send email
    GmailApp.sendEmail(recipientEmail, subject, "", {
      htmlBody: body.evaluate().getContent(),
      name: "The LaSallian Applications",
    });
  }

function takeHomeMay25Morning() {
takeHomeProcessSheet("May 25 Morning");
}

function takeHomeMay25Afternoon() {
takeHomeProcessSheet("May 25 Afternoon");
}

function takeHomeMay29Morning() {
takeHomeProcessSheet("May 29 Morning");
}

function takeHomeMay29Afternoon() {
    takeHomeProcessSheet("May 29 Afternoon");
}