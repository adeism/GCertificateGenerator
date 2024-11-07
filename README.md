# Google Apps Script - Certificate Generator 📜

This Google Apps Script automates generating personalized certificates, converting them to PDFs, and emailing them based on data submitted through a Google Form. It's designed for efficiency, ease of use, and robust error handling.

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
* **Google Form Integration:** Seamlessly works with Google Forms for user data input. 📝

## Prerequisites 📝

1. **Google Account:** You'll need a Google account to access Google Sheets, Forms, Slides, Drive, and Apps Script.
2. **Google Form:**
    * Create a new Google Form.
    * Add form fields that correspond to certificate placeholders (e.g., "Full Name", "Email", "Course Name"). **Field titles must match placeholders exactly (case-sensitive).**
    * In the **Responses** tab, click the spreadsheet icon and select "Create Spreadsheet".
3. **Google Slides Template:**
    * Create a Google Slides presentation as your certificate template.
    * Add placeholders (e.g., `{{Full Name}}`) where you want dynamic data.  **Placeholders must match form field titles exactly (case-sensitive).**
4. **Google Drive Folder:** Create a new folder in Google Drive to store generated PDFs. Note the folder's ID.
5. **Google Apps Script Editor (Bound to Sheet):**
    * Open the Google Sheet created by your Google Form.
    * Go to "Extensions" > "Apps Script".
6. **Script Configuration:**
    * At the top of the script:
        * `FOLDER_ID`: Replace `'YOUR_FOLDER_ID'` with your Drive folder's ID.
        * `TEMPLATE_ID`: Replace `'YOUR_TEMPLATE_ID'` with your Slides template's ID.
        * `EMAIL_SUBJECT`: Set the email subject.
        * `EMAIL_BODY`: Customize the email body (use placeholders).

## Example Email Output 📧

Here's an example of what the email sent to the recipient will look like:
   
     Subject: Your Certificate
   
     Dear John Doe,
   
     Please find your certificate attached.
   
     Sincerely,
     Your Organization

     Attachment: Certificate John Doe.pdf


## Setup ⚙️

1. **Open Script Editor:** Open the Google Sheet linked to your form. Go to "Extensions" > "Apps Script".
2. **Copy and Paste Code:** Copy and paste the provided code into the script editor.
3. **Save Script:** Save the script (File > Save or Ctrl/Cmd+S).
4. **Configure Constants:** Update the constants at the top of the script.
5. **Install Trigger:** In the Apps Script editor, go to "Edit" > "Current project's triggers". Click "+ Add Trigger" and configure:
    * `Function`: `onFormSubmit`
    * `Event source`: `From spreadsheet`
    * `Event type`: `On form submit`

## Usage 🚀

1. **Share Form:** Share your Google Form.
2. **Form Submission:** Users submit the form.
3. **Automatic Processing:** The script generates, converts, saves, and emails the certificate.
4. **Status and Logging:** The spreadsheet's "Status" column is updated, a PDF link is added, and logs are recorded.


## Example Data 📊 (in Google Form)

* **Full Name:** John Doe
* **Email:** john.doe@example.com
* **Course Name:** Introduction to Google Apps Script

## Troubleshooting ⚠️

* **Permissions:** Ensure the script has necessary permissions.
* **Error Log:** Check the "Error Log" sheet for errors.
* **Placeholders:** Verify placeholders match exactly.
* **Configuration:** Double-check the constants in the script.
