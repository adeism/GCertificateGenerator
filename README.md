# Google Apps Script - Certificate Generator 📜

This Google Apps Script automates the process of generating certificates based on data from a Google Sheet, converting them to PDFs, and emailing them to the specified recipients.  It's designed for efficiency and ease of use, with clear configuration and robust error handling.

## Features ✨

* **Automated Certificate Generation:** Creates personalized certificates from a Google Slides template. 📄
* **PDF Conversion:** Converts generated certificates into PDF format. PDF
* **Email Delivery:** Sends customized emails with the attached PDF certificate. 📧
* **Error Handling:** Logs errors to a dedicated "Error Log" sheet and updates the "Status" on the main sheet. ❌
* **Success Logging:** Records successful certificate generations in a "Log" sheet. ✅
* **Data Validation:** Validates email addresses and checks for required fields. ✔️
* **Dynamic Placeholders:** Uses placeholders in the template to dynamically populate certificate details. ✏️
* **Status Tracking:** Updates the status of each row in the spreadsheet (Processed, Failed). 🔄
* **PDF Link Storage:** Saves the link to the generated PDF in the spreadsheet for easy access. 🔗


## Prerequisites 📝

1. **Google Sheet:** Create a Google Sheet with columns for "Full Name", "Email", and any other data you want to include in the certificate. A "Status" column will be added automatically if it doesn't exist.
2. **Google Slides Template:** Create a Google Slides presentation as your certificate template.  Use placeholders like `{{Full Name}}`, `{{Email}}`, `{{Date}}` enclosed in double curly braces.
3. **Google Apps Script:** Open Script editor (Tools > Script editor) in your Google Sheet.
4. **Google Drive Folder:** Create a folder in Google Drive to store generated certificates.
5. **Configuration:**  At the top of the script, update the following constants:
    * `FOLDER_ID`:  The ID of your Google Drive folder.
    * `TEMPLATE_ID`: The ID of your Google Slides template.
    * `EMAIL_SUBJECT`:  The subject of the email.
    * `EMAIL_BODY`: The body of the email (use placeholders like `{{Full Name}}`).


## Setup ⚙️

1. Copy the provided Google Apps Script code into the Script editor.
2. Configure the constants at the top of the script with your specific IDs and email details.
3. **Install Triggers:** In the Apps Script editor (Edit > Current project's triggers), add a new trigger:
    * `Function`: `onFormSubmit`
    * `Event source`: `From spreadsheet`
    * `Event type`: `On form submit`

## Usage 🚀

1. Enter data into your Google Sheet, including "Full Name" and "Email" for each recipient.
2. Submitting the form (or directly editing the sheet) triggers the script.
3. The script generates a certificate, converts it to a PDF, saves it to your Google Drive folder, and emails it.
4. The sheet's "Status" column updates to "Processed" or "Failed", and a link to the PDF is added.
5. Logs are recorded in the "Log" and "Error Log" sheets.

## Example Data 📊

| Full Name | Email                 |
| --------- | --------------------- |
| John Doe  | john.doe@example.com   |
| Jane Doe  | jane.doe@example.com   |


## Troubleshooting ⚠️

* **Permissions:** Ensure the script has necessary permissions (Drive, Gmail, Slides).
* **Error Log:** Check the "Error Log" sheet for details on any errors.
* **Placeholders:** Verify placeholders in your Slides template match your sheet's column headers.
* **Configuration:** Double-check `FOLDER_ID`, `TEMPLATE_ID`, email subject, and body.
