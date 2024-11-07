# Bulk Email Sender using Google Apps Script

[![Licence](https://img.shields.io/github/license/Ileriayo/markdown-badges?style=for-the-badge)](./LICENSE)

## Support The Project

Your support is greatly appreciated and will help ensure all of the projects continued development and improvement. Thank you for being a part of the community!
You can send me money on Revolut by following this link: https://revolut.me/mnestorovv

## Overview

This Google Apps Script allows you to send bulk emails based on data from a Google Sheets document. The script reads the email template and recipient details from the Google Sheets, sends emails accordingly, and logs the email status in another sheet.

## Features

- **Bulk Email Sending**: Sends emails to multiple recipients with customized content.
- **Template-Based Emails**: Email content can be customized using templates.
- **Dynamic Content Replacement**: Order ID, COD Price, Currency, and Name placeholders in the email body and subject are replaced with actual values.
   - **Dynamic Tags**: `<<Order ID>>` | `<<COD Price>>` | `<<Name>>` | `<<Currency>>` | `<<PAYMENT_DETAILS>>`
- **Logging**: Logs the status of each email sent, including the timestamp and the sender's email address.
- **Fallback Email Handling**: Uses a default sender email if none is provided in the template.

## Setup

1. **Create a Google Sheet**:

   - Create a Google Sheet with three sheets: `Orders`, `Email Template` and `Email Log`.
   
2. **Orders Sheet**:

   - The `Orders` sheet should have the following columns starting from the first row:

     ```
     | Order ID | COD Price | Currency | Name | Email | Language |
     ```
   - Fill in the order details accordingly.

3. **Email Template Sheet**:

   - The `Email Template` sheet should have the following structure:

     ```
     | Content       | en            | es            | fr            | ... |
     | --------------|---------------|---------------|---------------|-----|
     | Email Title   | Title in EN   | Title in ES   | Title in FR   | ... |
     | Email Subject | Subject in EN | Subject in ES | Subject in FR | ... |
     | Email Header  | Header in EN  | Header in ES  | Header in FR  | ... |
     | Email Body    | Body in EN    | Body in ES    | Body in FR    | ... |
     | Sender Email  | sender@example.com                                  |
     | Logo URL      | https://example.com/logo.png                        |
     -----------------------------------------------------------------------
     ```

   - Example `Email Subject` template 1:

     ```
     SITE_NAME order number #<<Order ID>> - Order received
     ```
    
   - Example `Email Body` template 1:

     ```
     Hello <<Name>>, <br><br>

     Your Order ID is <<Order ID>> with a price of <<COD Price>> <<Currency>>. <br><br>

     Thank you for your purchase! <br><br>

     Regards,<br>
     SITE_NAME!
     ```

    - Example `Email Subject` template 2:

      ```
      SITE_NAME order number #<<Order ID>> - Missing payment
      ```
    
   - Example `Email Body` template 2:

      ```
      Hello <<Name>>,<br><br> 

      We are contacting you regarding your SITE_NAME order <b>#<<Order ID>></b> and amount to pay: <b><<COD Price>> <<Currency>></b>.<br><br> 
      
      We have not yet received payment for your order.<br><br> 
      
      You can pay for your order via: <br><br> 
      
      <<PAYMENT_DETAILS>>
      
      As a reason for payment, enter your order ID.<br><br> 
      
      Thank you in advance and we apologize for the inconvenience!<br><br> 
      
      Sincerely,<br>
      SITE_NAME!
      ```    

4. **Payment Details Sheet**

   - The Payment Details sheet should have the following structure:
   
     ```
     | en           |                ...                |          ...          |
     ----------------------------------------------------------------------------
     |              | PayPal                            | your_email@address    |
     |              |                                   |                       |
     |              | You can make payment also via:    |                       |
     |              | Beneficient                       | YOUR COMPANY          |
     |              | Bank                              | YOUR BANK             |
     |              | IBAN                              | YOUR IBAN             |
     |              | SWIFT                             | BANK SWIFT            |
     |              | Bank Address                      | BANK ADDRESS          |
     ----------------------------------------------------------------------------
     ```

6. **Email Log Sheet**

   - The Email Log sheet should have the following structure:

     ```
     | Timestamp | Order ID | COD Price | Currency | Name | Email | Language | Status | Send By | SM |
     ```
   - `SM` is the count of the sended emails for each user.
     
7. **Google Apps Script**:

   - Open the script editor in your Google Sheet (`Extensions` > `Apps Script`).
   - Copy and paste the provided script into the script editor.
   - Save the script.

8. **Authorization**:

   - Run the script for the first time and authorize the required permissions.

## Usage

1. **Run the Script**:

   - Open the Google Sheet and run the script by clicking the button or running the function `sendBulkEmails` from the script editor.
   - A confirmation dialog will appear. Click "Yes" to send the emails.

3. **Check Logs**:

   - After running the script, check the `Email Log` sheet to see the log entries, including the timestamp, order details, status, and sender email.

## Code.gs

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
    if (!sheet) {
      Logger.log('Orders sheet not found');
      return;
    }
    var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6); // Read 6 columns: Order ID, COD Price, Currency, Name, Email, Language
    var data = dataRange.getValues();
  
    var templateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email Template');
    if (!templateSheet) {
      Logger.log('Email Template sheet not found');
      return;
    }
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
    var paymentDetails = getPaymentDetails(); // Get payment details from the Payment Details sheet
  
    for (var i = 0; i < data.length; i++) {
      var orderId = data[i][0];
      var codPrice = data[i][1];
      var currency = data[i][2];
      var name = data[i][3];
      var email = data[i][4];
      var language = data[i][5] || 'en'; // Default to English if language is empty
      
      var template = templates[language] || templates['en']; // Default to English if language not found

      var emailSubject = template.emailSubject.replace('<<Order ID>>', orderId);
      var emailBody = template.emailBodyTemplate.replace('<<Order ID>>', orderId)
                                                 .replace('<<COD Price>>', codPrice)
                                                 .replace('<<Name>>', name)
                                                 .replace('<<Currency>>', currency);

      var paymentInfo = paymentDetails[language] || {};
      var paymentDetailsString = formatPaymentDetails(paymentInfo);

      emailBody = emailBody.replace('<<PAYMENT_DETAILS>>', paymentDetailsString);

      emailBody = `
        <div style="max-width: 600px; margin: 0 auto; border: 1px solid #fff; padding: 0; font-family: Arial, sans-serif;">
          <div style="text-align: center;">
            <img src="${logoUrl}" alt="Logo" style="max-width: 140px; height: auto;"/>
          </div>
          <div style="background-color: #e91e63; color: white; text-align: center; padding: 10px; font-size: 24px; font-weight: bold; text-transform: capitalize;">
            ${template.emailHeader}
          </div>
          <div style="padding: 20px; font-size: 16px; color: #333;">
            ${emailBody}
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

function getPaymentDetails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payment Details');
  if (!sheet) {
    Logger.log('Payment Details sheet not found');
    return {};
  }
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  var paymentDetails = {};
  var currentCountry = null;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row[0] && !row[1]) { // New country section
      currentCountry = row[0];
      paymentDetails[currentCountry] = [];
    } else if (currentCountry && row[1] === "" && row[2] === "") { // Handle blank line
      paymentDetails[currentCountry].push({label: "", value: ""});
    } else if (currentCountry && row[1] && row[2]) { // Add details to current country
      paymentDetails[currentCountry].push({label: row[1], value: row[2]});
    }
  }

  return paymentDetails;
}

function formatPaymentDetails(paymentInfo) {
  var paymentDetailsString = "";
  for (var i = 0; i < paymentInfo.length; i++) {
    if (paymentInfo[i].label === "" && paymentInfo[i].value === "") {
      paymentDetailsString += `<br>`; // Add blank line
    } else {
      paymentDetailsString += `<b>${paymentInfo[i].label}:</b> ${paymentInfo[i].value}<br>`;
    }
  }
  return paymentDetailsString;
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
    logSheet.appendRow(['Timestamp', 'Order ID', 'COD Price', 'Currency', 'Name', 'Email', 'Language', 'Status', 'Sent From', 'SE']);
  }
  return logSheet;
}

function logEmailResult(logSheet, orderId, codPrice, currency, name, email, language, status, senderEmail) {
  var timestamp = new Date();
  var formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var sentEmailCount = getEmailCount(logSheet, email) + 1; // Increment the count by 1
  logSheet.appendRow([formattedTimestamp, orderId, codPrice, currency, name, email, language, status, senderEmail, sentEmailCount]);
}

function getEmailCount(logSheet, email) {
  var data = logSheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][5] === email) {
      count++;
    }
  }
  return count;
}

function logCancelStatus() {
  var logSheet = getLogSheet();
  var timestamp = new Date();
  var formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var senderEmail = getUserEmail(); // Get the email of the person running the script
  logSheet.appendRow([formattedTimestamp, '', '', '', '', '', '', 'Canceled', senderEmail, '']);
}
```
## appsscript.json

```json
{
  "timeZone": "Europe/Sofia",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "Gmail",
        "version": "v1",
        "serviceId": "gmail"
      },
      {
        "userSymbol": "Drive",
        "version": "v3",
        "serviceId": "drive"
      },
      {
        "userSymbol": "Sheets",
        "version": "v4",
        "serviceId": "sheets"
      }
    ]
  },
  "oauthScopes": [
    "https://mail.google.com/",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/script.external_request"
  ],
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

## Contributing

We welcome contributions from the community! If you would like to contribute, please fork the repository, make changes, and submit a pull request. You can also open an issue to discuss potential changes.

## Thank You

If you have any questions, feel free to reach out.

## Support The Project

If you find this script helpful and would like to support its development and maintenance, please consider the following options:

- **_Star the repository_**: If you're using this script from a GitHub repository, please give the project a star on GitHub. This helps others discover the project and shows your appreciation for the work done.

- **_Share your feedback_**: Your feedback, suggestions, and feature requests are invaluable to the project's growth. Please open issues on the GitHub repository or contact the author directly to provide your input.

- **_Contribute_**: You can contribute to the project by submitting pull requests with bug fixes, improvements, or new features. Make sure to follow the project's coding style and guidelines when making changes.

- **_Spread the word_**: Share the project with your friends, colleagues, and social media networks to help others benefit from the script as well.

- **_Donate_**: Show your appreciation with a small donation. Your support will help me maintain and enhance the script. Every little bit helps, and your donation will make a big difference in my ability to keep this project alive and thriving.

Your support is greatly appreciated and will help ensure all of the projects continued development and improvement. Thank you for being a part of the community!
You can send me money on Revolut by following this link: https://revolut.me/mnestorovv

---

## License

This project is released under the MIT License.
