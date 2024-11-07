/**
 * Customizable Settings - Modify these values as needed
 */
const FOLDER_ID = 'YOUR_FOLDER_ID'; // Google Drive folder ID where certificates will be saved
const TEMPLATE_ID = 'YOUR_TEMPLATE_ID'; // Google Slides template ID for the certificate
const EMAIL_SUBJECT = 'Your Certificate'; // Subject line of the email sent to the user
const EMAIL_BODY = `Dear {{Full Name}},\n\nPlease find your certificate attached.\n\nSincerely,\nYour Organization`; // Email body with placeholders for personalization

/**
 * Main function triggered on form submission
 * @param {Object} e - Event object from the form submission
 */
function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.range.getSheet();
  const rowIndex = e.range.getRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; // Retrieve header row for mapping

  try {
    const rowData = getRowData(headers, e.values); // Extract data from the submitted row
    const statusColIndex = getStatusColumnIndex(sheet, headers); // Get or create 'Status' column index

    // Skip processing if already marked as 'Processed'
    if (rowData['Status'] === 'Processed') return;

    // Validate data
    if (!isValidData(rowData)) {
      logError(ss, rowIndex, rowData, 'Incomplete data or invalid email');
      setStatus(sheet, rowIndex, statusColIndex, 'Failed'); // Mark as failed if data is invalid
      return;
    }

    // Generate the certificate PDF
    const pdfBlob = generateCertificate(rowData);

    // Send certificate email to the user
    sendEmail(rowData, pdfBlob);

  } catch (error) {
    // Log errors that occur during processing
    logError(ss, rowIndex, rowData, `Main error: ${error.message}`);
    setStatus(sheet, rowIndex, headers.indexOf('Status') + 1, 'Failed'); // Mark as 'Failed' if error occurs
  }
}

/**
 * Helper functions - These handle specific tasks within the main function
 */

function getRowData(headers, values) {
  // Maps form submission values to headers, creating a key-value object of row data
  const rowData = {};
  headers.forEach((header, index) => rowData[header] = values[index]);
  return rowData;
}

function getStatusColumnIndex(sheet, headers) {
  // Finds or adds 'Status' column to track processing status
  let statusColIndex = headers.indexOf('Status');
  if (statusColIndex === -1) {
    statusColIndex = sheet.getLastColumn();
    sheet.insertColumnAfter(statusColIndex); // Add new column if not found
    sheet.getRange(1, statusColIndex + 1).setValue('Status');
  }
  return statusColIndex;
}

function isValidData(rowData) {
  // Validates that required fields (name and email) are present and that the email format is correct
  const name = rowData['Full Name'];
  const email = rowData['Email'];
  return name && email && validateEmail(email);
}

function generateCertificate(rowData) {
  // Generates a certificate using Google Slides template, replaces placeholders, and converts it to PDF
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const slideCopy = DriveApp.getFileById(TEMPLATE_ID).makeCopy(`Certificate ${rowData['Full Name']}`, folder);
  const presentation = SlidesApp.openById(slideCopy.getId());

  replacePlaceholders(presentation, rowData); // Fill in placeholder text with user data
  presentation.saveAndClose();

  const pdfBlob = slideCopy.getAs(MimeType.PDF); // Convert to PDF
  const pdfFile = folder.createFile(pdfBlob).setName(`Certificate ${rowData['Full Name']}.pdf`);
  slideCopy.setTrashed(true); // Delete slide copy after creating PDF

  // Update sheet with a link to the certificate PDF
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowIndex = sheet.getActiveRange().getRow();
  const statusColIndex = headers.indexOf('Status');
  sheet.getRange(rowIndex, statusColIndex + 2).setValue(pdfFile.getUrl());

  return pdfBlob;
}

function replacePlaceholders(presentation, rowData) {
  // Replaces placeholder text on each slide with values from rowData
  const currentDate = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  presentation.getSlides().forEach(slide => {
    for (const key in rowData) {
      slide.replaceAllText(`{{${key}}}`, rowData[key] || '');
    }
    slide.replaceAllText('{{Date}}', currentDate); // Insert current date
  });
}

function sendEmail(rowData, pdfBlob) {
  // Sends the generated PDF certificate via email to the user
  let emailMessage = EMAIL_BODY;
  const currentDate = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });

  for (const key in rowData) {
    emailMessage = emailMessage.replace(`{{${key}}}`, rowData[key] || '');
  }
  emailMessage = emailMessage.replace('{{Date}}', currentDate); // Insert current date in email body

  try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getActiveSheet();
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const rowIndex = sheet.getActiveRange().getRow();
      const statusColIndex = headers.indexOf('Status');

      GmailApp.sendEmail(rowData['Email'], EMAIL_SUBJECT, emailMessage, { attachments: [pdfBlob] });
      setStatus(sheet, rowIndex, statusColIndex, 'Processed'); // Mark as 'Processed' upon success
      logSuccess(ss, rowIndex, rowData, pdfBlob.getUrl());
  } catch (emailError) {
    // Handle email errors by logging and updating the status to 'Failed'
    logError(ss, rowIndex, rowData, `Failed to send email: ${emailError.message}`);
    setStatus(sheet, rowIndex, headers.indexOf('Status')+1, 'Failed');
  }
}

function validateEmail(email) {
  // Regex to validate email format
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function setStatus(sheet, rowIndex, statusColIndex, status) {
  // Sets the status cell value for the current row
  sheet.getRange(rowIndex, statusColIndex + 1).setValue(status);
}

function logSuccess(ss, rowIndex, rowData, certLink) {
  // Logs successful certificate creation and email sending
  logToSheet(ss, 'Log', ['Timestamp', 'Full Name', 'Email', 'Certificate Link', 'Status'], [new Date(), rowData['Full Name'], rowData['Email'], certLink, 'Success']);
}

function logError(ss, rowIndex, rowData, errorMessage) {
  // Logs errors that occur during the process
  const sheet = ss.getActiveSheet();
  setStatus(sheet, rowIndex, sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('Status'), 'Failed');
  logToSheet(ss, 'Error Log', ['Timestamp', 'Full Name', 'Email', 'Error'], [new Date(), rowData['Full Name'] || 'N/A', rowData['Email'] || 'N/A', errorMessage]);
}

function logToSheet(ss, sheetName, headers, data) {
  // Appends log data to specified log sheet, creating it if necessary
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
  }
  sheet.appendRow(data);
}
