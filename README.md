# Bulk Email Sender using Google Apps Script

[![Licence](https://img.shields.io/github/license/Ileriayo/markdown-badges?style=for-the-badge)](./LICENSE)

## Overview

This Google Apps Script allows you to send bulk emails based on data from a Google Sheets document. The script reads the email template and recipient details from the Google Sheets, sends emails accordingly, and logs the email status in another sheet.

## Features

- **Bulk Email Sending**: Sends emails to multiple recipients with customized content.
- **Template-Based Emails**: Email content can be customized using templates.
- **Dynamic Content Replacement**: Order ID, COD Price, Currency, and Name placeholders in the email body and subject are replaced with actual values.
- **Logging**: Logs the status of each email sent, including the timestamp and the sender's email address.
- **Fallback Email Handling**: Uses a default sender email if none is provided in the template.

## Setup

1. **Create a Google Sheet**:

   - Create a Google Sheet with two sheets: `Orders` and `Email Template`.
   
3. **Orders Sheet**:

   - The `Orders` sheet should have the following columns starting from the first row:

     ```
     | Order ID | COD Price | Currency | Name | Email | Language |
     ```
   - Fill in the order details accordingly.

4. **Email Template Sheet**:

   - The `Email Template` sheet should have the following structure:

     ```
     | Content       | en            | es            | fr            | ... |
     | --------------|---------------|---------------|---------------|-----|
     | Email Title   | Title in EN   | Title in ES   | Title in FR   | ... |
     | Email Subject | Subject in EN | Subject in ES | Subject in FR | ... |
     | Email Header  | Header in EN  | Header in ES  | Header in FR  | ... |
     | Email Body    | Body in EN    | Body in ES    | Body in FR    | ... |
     | Sender Email  | sender@example.com
     | Logo URL      | https://example.com/logo.png
     ```

   - Example `Email Subject` text:

     ```
     SITE_NAME order number #<<Order ID>> - Missing payment
     ```
    
   - Example `Email Body` text:

     ```
     Hello <<Name>>, <br><br> Your Order ID is <<Order ID>> with a COD price of <<COD Price>> <<Currency>>. <br><br> Thank you for your purchase! <br><br> Regards,<br>SITE_NAME!
     ```

6. **Google Apps Script**:

   - Open the script editor in your Google Sheet (`Extensions` > `Apps Script`).
   - Copy and paste the provided script into the script editor.
   - Save the script.

7. **Authorization**:

   - Run the script for the first time and authorize the required permissions.

## Usage

1. **Run the Script**:

   - Open the Google Sheet and run the script by clicking the button or running the function `sendBulkEmails` from the script editor.
   - A confirmation dialog will appear. Click "Yes" to send the emails.

3. **Check Logs**:

   - After running the script, check the `Email Log` sheet to see the log entries, including the timestamp, order details, status, and sender email.

## Script

```javascript
function sendBulkEmails() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Send Emails', 'Do you want to send emails to all users?', ui.ButtonSet.YES_NO);

  if (response == ui.Button.NO) {
    Logger.log('User canceled the email sending process.');
    logCancelStatus();
    return;
  }

  try {
    var userEmail = getUserEmail(); // Get the email of the person running the script

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Orders');
    var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6); // Read 6 columns: Order ID, COD Price, Currency, Name, Email, Language
    var data = dataRange.getValues();
  
    var templateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email Template');
    var templateHeaders = templateSheet.getRange(1, 2, 1, templateSheet.getLastColumn() - 1).getValues()[0]; // Get language headers
    var templates = {};

    // Common values for all languages
    var senderEmail = templateSheet.getRange('B6').getValue(); // Read sender email
    var logoUrl = templateSheet.getRange('B7').getValue(); // Read logo URL

    // Use user email if sender email is empty
    if (!senderEmail) {
      senderEmail = userEmail;
    }

    // Load templates for each language using ISO codes
    for (var i = 0; i < templateHeaders.length; i++) {
      var isoCode = templateHeaders[i];
      templates[isoCode] = {
        emailTitle: templateSheet.getRange(2, i + 2).getValue(),
        emailSubject: templateSheet.getRange(3, i + 2).getValue(),
        emailHeader: templateSheet.getRange(4, i + 2).getValue(),
        emailBodyTemplate: templateSheet.getRange(5, i + 2).getValue()
      };
    }

    var logSheet = getLogSheet();
  
    for (var i = 0; i < data.length; i++) {
      var orderId = data[i][0];
      var codPrice = data[i][1];
      var currency = data[i][2];
      var name = data[i][3];
      var email = data[i][4];
      var language = data[i][5];
      
      var template = templates[language] || templates['en']; // Default to English if language not found
  
      var emailSubject = template.emailSubject.replace('<<Order ID>>', orderId);
      var emailBody = `
        <div style="max-width: 600px; margin: 0 auto; border: 1px solid #fff; padding: 0; font-family: Arial, sans-serif;">
          <div style="text-align: center;">
            <img src="${logoUrl}" alt="Logo" style="max-width: 140px; height: auto;"/>
          </div>
          <div style="background-color: #e91e63; color: white; text-align: center; padding: 10px; font-size: 24px; font-weight: bold; text-transform: capitalize;">
            ${template.emailHeader}
          </div>
          <div style="padding: 20px; font-size: 16px; color: #333;">
            ${template.emailBodyTemplate.replace('<<Order ID>>', orderId)
                                        .replace('<<COD Price>>', codPrice)
                                        .replace('<<Name>>', name)
                                        .replace('<<Currency>>', currency)}
          </div>
        </div>
      `;
  
      try {
        GmailApp.sendEmail(
          email,
          emailSubject, // Use the dynamic subject with Order ID
          '', // Plain text body is empty because we use HTML body
          {
            from: senderEmail, // Use the sender email if provided, otherwise use the user's email
            name: template.emailTitle, // Set the name to appear in the sender field,
            htmlBody: emailBody // HTML body content
          }
        );
        Logger.log('Email sent to: ' + email);
        logEmailResult(logSheet, orderId, codPrice, currency, name, email, language, 'Sent', senderEmail);
      } catch (e) {
        Logger.log('Failed to send email to: ' + email + '. Error: ' + e.message);
        logEmailResult(logSheet, orderId, codPrice, currency, name, email, language, 'Failed', senderEmail);
      }
  
      Utilities.sleep(1000); // 1 second delay between emails
    }
  } catch (e) {
    Logger.log('Error in sendBulkEmails function: ' + e.message);
  }
}

function getUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (e) {
    Logger.log('Error getting user email: ' + e.message);
    return 'Unknown';
  }
}

function getLogSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = spreadsheet.getSheetByName('Email Log');
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet('Email Log');
  }
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(['Timestamp', 'Order ID', 'COD Price', 'Currency', 'Name', 'Email', 'Language', 'Status', 'Sent From']);
  }
  return logSheet;
}

function logEmailResult(logSheet, orderId, codPrice, currency, name, email, language, status, senderEmail) {
  var timestamp = new Date();
  var formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  logSheet.appendRow([formattedTimestamp, orderId, codPrice, currency, name, email, language, status, senderEmail]);
}

function logCancelStatus() {
  var logSheet = getLogSheet();
  var timestamp = new Date();
  var formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var senderEmail = getUserEmail(); // Get the email of the person running the script
  logSheet.appendRow([formattedTimestamp, '', '', '', '', '', '', 'Canceled', senderEmail]);
}
```
## Contributing

We welcome contributions from the community! If you would like to contribute, please fork the repository, make changes, and submit a pull request. You can also open an issue to discuss potential changes.

## Thank You

If you have any questions, feel free to reach out.

---

## License

This project is released under the MIT License.
