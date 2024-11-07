/**
 * Customizable Settings - Modify these values as needed
 */
const FOLDER_ID = 'YOUR_FOLDER_ID'; // Replace with your Google Drive folder ID
const TEMPLATE_ID = 'YOUR_TEMPLATE_ID'; // Replace with your Google Slides template ID
const EMAIL_SUBJECT = 'Your Certificate'; // Subject of the email
const EMAIL_BODY = `Dear {{Full Name}},\n\nPlease find your certificate attached.\n\nSincerely,\nYour Organization`; // Body of the email


/**
 * Main function triggered on form submission
 * @param {Object} e - Form submission event object
 */
function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.range.getSheet();
  const rowIndex = e.range.getRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  try {
    const rowData = getRowData(headers, e.values);
    const statusColIndex = getStatusColumnIndex(sheet, headers);

    // Check if already processed
    if (rowData['Status'] === 'Processed') return;

    // Data validation
    if (!isValidData(rowData)) {
      logError(ss, rowIndex, rowData, 'Incomplete data or invalid email');
      setStatus(sheet, rowIndex, statusColIndex, 'Failed');
      return;
    }

    // Generate certificate
    const pdfBlob = generateCertificate(rowData);

    // Send email
    sendEmail(rowData, pdfBlob);

  } catch (error) {
    logError(ss, rowIndex, rowData, `Main error: ${error.message}`);
    setStatus(sheet, rowIndex, headers.indexOf('Status') + 1, 'Failed'); // Use index + 1 for setting value.
  }
}


/**
 * Helper functions - These handle specific tasks within the main function
 */

function getRowData(headers, values) {
  const rowData = {};
  headers.forEach((header, index) => rowData[header] = values[index]);
  return rowData;
}

function getStatusColumnIndex(sheet, headers) {
  let statusColIndex = headers.indexOf('Status');
  if (statusColIndex === -1) {
    statusColIndex = sheet.getLastColumn();
    sheet.insertColumnAfter(statusColIndex);
    sheet.getRange(1, statusColIndex + 1).setValue('Status');
  }
  return statusColIndex;
}


function isValidData(rowData) {
  const name = rowData['Full Name'];
  const email = rowData['Email'];
  return name && email && validateEmail(email);
}

function generateCertificate(rowData) {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const slideCopy = DriveApp.getFileById(TEMPLATE_ID).makeCopy(`Certificate ${rowData['Full Name']}`, folder);
  const presentation = SlidesApp.openById(slideCopy.getId());

  replacePlaceholders(presentation, rowData);
  presentation.saveAndClose();

  const pdfBlob = slideCopy.getAs(MimeType.PDF);
  const pdfFile = folder.createFile(pdfBlob).setName(`Certificate ${rowData['Full Name']}.pdf`);
  slideCopy.setTrashed(true);

  // Add PDF link to sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowIndex = sheet.getActiveRange().getRow(); // Changed to get active row since we're in another function
  const statusColIndex = headers.indexOf('Status');
  sheet.getRange(rowIndex, statusColIndex + 2).setValue(pdfFile.getUrl());


  return pdfBlob;
}



function replacePlaceholders(presentation, rowData) {
  const currentDate = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  presentation.getSlides().forEach(slide => {
    for (const key in rowData) {
      slide.replaceAllText(`{{${key}}}`, rowData[key] || '');
    }
    slide.replaceAllText('{{Date}}', currentDate);
  });
}



function sendEmail(rowData, pdfBlob) {
  let emailMessage = EMAIL_BODY;
  const currentDate = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });


  for (const key in rowData) {
    emailMessage = emailMessage.replace(`{{${key}}}`, rowData[key] || '');
  }
  emailMessage = emailMessage.replace('{{Date}}', currentDate);

  try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getActiveSheet();
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const rowIndex = sheet.getActiveRange().getRow();
      const statusColIndex = headers.indexOf('Status');

      GmailApp.sendEmail(rowData['Email'], EMAIL_SUBJECT, emailMessage, { attachments: [pdfBlob] });
      setStatus(sheet, rowIndex, statusColIndex, 'Processed');
      logSuccess(ss, rowIndex, rowData, pdfBlob.getUrl());
  } catch (emailError) {
    logError(ss, rowIndex, rowData, `Failed to send email: ${emailError.message}`);
    setStatus(sheet, rowIndex, headers.indexOf('Status')+1, 'Failed');

  }
}



function validateEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}


function setStatus(sheet, rowIndex, statusColIndex, status) {
  sheet.getRange(rowIndex, statusColIndex + 1).setValue(status);
}


function logSuccess(ss, rowIndex, rowData, certLink) {
  logToSheet(ss, 'Log', ['Timestamp', 'Full Name', 'Email', 'Certificate Link', 'Status'], [new Date(), rowData['Full Name'], rowData['Email'], certLink, 'Success']);
}

function logError(ss, rowIndex, rowData, errorMessage) {
    const sheet = ss.getActiveSheet();
    setStatus(sheet, rowIndex, sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('Status'), 'Failed');

  logToSheet(ss, 'Error Log', ['Timestamp', 'Full Name', 'Email', 'Error'], [new Date(), rowData['Full Name'] || 'N/A', rowData['Email'] || 'N/A', errorMessage]);
}


function logToSheet(ss, sheetName, headers, data) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
  }
  sheet.appendRow(data);
}
