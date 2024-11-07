# Google Apps Script - Certificate Generator 📜

This Google Apps Script automatically generates certificates based on data from a Google Sheet, converts them to PDFs, and emails them to the recipients.

## Features ✨

* **Automated Certificate Generation:** Creates certificates from a Google Slides template. 📄
* **PDF Conversion:** Converts generated certificates to PDF format.  PDF
* **Email Delivery:** Sends personalized emails with the attached PDF certificate. 📧
* **Error Handling:** Logs errors to a dedicated error log sheet and updates the status on the main sheet. ❌
* **Success Logging:** Logs successful certificate generations to a dedicated log sheet. ✅
* **Data Validation:** Validates email addresses and checks for required fields. ✔️
* **Dynamic Placeholders:** Uses placeholders in the template to populate certificate details. ✏️
* **Status Tracking:** Updates the status of each row in the spreadsheet (Processed, Failed). 🔄
* **PDF Link Storage:** Stores the link to the generated PDF in the spreadsheet. 🔗


## Prerequisites 📝

1. **Google Sheet:** Create a Google Sheet with the required columns: "Full Name", "Email", and any other data you want to include in the certificate.  A "Status" column will be added automatically if it doesn't exist.
2. **Google Slides Template:** Create a Google Slides presentation that serves as your certificate template. Use placeholders enclosed in double curly braces `{{PlaceholderName}}` for dynamic content.  For example, `{{Full Name}}`, `{{Email}}`, `{{Date}}`.
3. **Google Apps Script:** Open Script editor (Tools > Script editor) in your Google Sheet.
4. **Folder ID:**  Obtain the ID of the Google Drive folder where you want to save the generated certificates and PDFs. Replace `FOLDER_ID` in the script with your folder ID.
5. **Template ID:** Obtain the ID of your Google Slides template file. Replace `TEMPLATE_ID` in the script with your template ID.
6. **Email Settings:**  
    * Customize `EMAIL_SUBJECT` and `EMAIL_BODY` variables with your desired subject and email body. Use placeholders like `{{Full Name}}` to personalize the email.
    * Ensure that the script has permission to send emails on your behalf.

## Setup ⚙️

1. Copy the provided Google Apps Script code into your Script editor.
2. Replace the placeholder values for `FOLDER_ID`, `TEMPLATE_ID`, `EMAIL_SUBJECT`, and `EMAIL_BODY` with your actual values.
3. **Install Triggers:** In the Apps Script editor, go to "Edit" > "Current project's triggers".  Add a new trigger with the following settings:
    * `Choose which function to run`: `onFormSubmit`
    * `Select event source`: `From spreadsheet`
    * `Select event type`: `On form submit`

## Usage 🚀

1. Enter data into your Google Sheet, including the "Full Name" and "Email" for each recipient. Submitting the form (either via a form linked to the sheet, or by directly editing the sheet if the trigger is set to "On form submit") will trigger the script.
2. The script will generate a certificate, convert it to PDF, save it to the specified Google Drive folder, and email it to the recipient.
3. The "Status" column in the sheet will be updated to "Processed" or "Failed" to indicate the outcome. A link to the PDF will also be added to a column next to the "Status" column.
4. Logs of successful operations and errors will be recorded in the "Log" and "Error Log" sheets respectively.


## Example Data 📊

| Full Name | Email                 |
| --------- | --------------------- |
| John Doe  | john.doe@example.com   |
| Jane Doe  | jane.doe@example.com   |


## Troubleshooting ⚠️

* **Permissions:** Ensure the script has the necessary permissions (Drive, Gmail, Slides). You'll be prompted to authorize these permissions when you run the script for the first time.
* **Error Log:** Check the "Error Log" sheet for any errors encountered during the process.
* **Placeholders:** Double-check that the placeholders in your Google Slides template match the column headers in your Google Sheet.
* **Folder/Template IDs:** Verify that the `FOLDER_ID` and `TEMPLATE_ID` are correct.

## License

This project is open-source and available under the [MIT License](LICENSE). Feel free to modify and use it according to your needs. <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="MIT License">
