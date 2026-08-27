/**
 * IRCA Google Apps Script backend
 * ----------------------------------------------------------
 * Handles submissions from both student-registration.html
 * and the new donation page.
 */

const DRIVE_FOLDER_NAME = 'IRCA Admission Uploads';

// ---- Email notification settings ----
const ADMIN_EMAIL = 'irca.admin@gmail.com,riyadeb.work@gmail.com,dipsraj.kundu@gmail.com';
const SEND_ADMIN_EMAIL = true;
const SEND_APPLICANT_CONFIRMATION = true;

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const folder = getOrCreateFolder(DRIVE_FOLDER_NAME);

    // Route the submission based on the payload data
    // If donorName exists, it's from the donation page. Otherwise, it's an admission.
    if (data.donorName !== undefined) {
      return processDonation(data, folder);
    } else {
      return processAdmission(data, folder);
    }

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 1. ADMISSION FORM HANDLER
// ==========================================
function processAdmission(data, folder) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Submissions') || createSubmissionsSheet();

  const photoUrl = saveFile(folder, data.photoFileName, data.photoBase64, data.referenceNo, 'photo');
  const paymentUrl = saveFile(folder, data.paymentFileName, data.paymentBase64, data.referenceNo, 'payment');

  sheet.appendRow([
    data.submittedAt,
    data.referenceNo,
    data.candidateName,
    data.guardianName,
    data.address,
    data.email,
    data.phone,
    data.whatsapp,
    data.dob,
    data.gender,
    data.education,
    data.aadhar,
    data.profession,
    data.subjects,
    data.amountPayable,
    data.utr,
    photoUrl,
    paymentUrl
  ]);

  sendAdmissionNotificationEmails(data, photoUrl, paymentUrl);

  return ContentService.createTextOutput(JSON.stringify({ result: 'success', referenceNo: data.referenceNo }))
      .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 2. DONATION FORM HANDLER
// ==========================================
function processDonation(data, folder) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('donation') || createDonationsSheet();

  // Donations don't generate a referenceNo on the frontend, so we use the phone number to name the file
  const filePrefix = data.donorPhone || new Date().getTime();
  const paymentUrl = saveFile(folder, data.paymentFileName, data.paymentBase64, filePrefix, 'donation_receipt');

  sheet.appendRow([
    data.submittedAt,
    data.donorName,
    data.donorPhone,
    data.donorEmail,
    data.donorAmount,
    data.donorUtr,
    data.donorMessage,
    data.publicListingConsent ? 'Yes' : 'No',
    paymentUrl
  ]);

  sendDonationNotificationEmail(data, paymentUrl);

  return ContentService.createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// EMAIL NOTIFICATIONS
// ==========================================
function sendAdmissionNotificationEmails(data, photoUrl, paymentUrl) {
  const summary = [
    'Reference No: ' + data.referenceNo,
    'Candidate Name: ' + data.candidateName,
    'Guardian Name: ' + data.guardianName,
    'Address: ' + data.address,
    'Email: ' + data.email,
    'Phone: ' + data.phone,
    'WhatsApp: ' + data.whatsapp,
    'DOB: ' + data.dob,
    'Gender: ' + data.gender,
    'Education: ' + data.education,
    'Aadhar: ' + data.aadhar,
    'Profession: ' + data.profession,
    'Subjects: ' + data.subjects,
    'Amount Payable: ' + data.amountPayable,
    'UTR: ' + data.utr,
    'Photo: ' + (photoUrl || 'not uploaded'),
    'Payment Screenshot: ' + (paymentUrl || 'not uploaded')
  ].join('\n');

  if (SEND_ADMIN_EMAIL && ADMIN_EMAIL) {
    try {
      MailApp.sendEmail({
        to: ADMIN_EMAIL,
        subject: 'IRCA - New Student Registration - ' + data.candidateName + ' (' + data.referenceNo + ')',
        body: 'A new admission form was submitted.\n\n' + summary
      });
    } catch (err) {}
  }

  if (SEND_APPLICANT_CONFIRMATION && data.email) {
    try {
      MailApp.sendEmail({
        to: data.email,
        subject: 'IRCA Admission Received — Reference ' + data.referenceNo,
        body: 'Dear ' + data.candidateName + ',\n\n' +
            'Thank you for applying to Ichapur Rupsa Cultural Association. We have received your admission form.\n\n' +
            'Your reference number is: ' + data.referenceNo + '\n' +
            'Subjects applied for: ' + data.subjects + '\n' +
            'Amount payable: ' + data.amountPayable + '\n\n' +
            'Our office will verify your payment and confirm your seat shortly. Please quote your reference number in any follow-up.\n\n' +
            'Regards,\nIchapur Rupsa Cultural Association'
      });
    } catch (err) {}
  }
}

function sendDonationNotificationEmail(data, paymentUrl) {
  if (!SEND_ADMIN_EMAIL || !ADMIN_EMAIL) return;

  const summary = [
    'Donor Name: ' + data.donorName,
    'Phone: ' + data.donorPhone,
    'Email: ' + (data.donorEmail || 'N/A'),
    'Amount: ' + (data.donorAmount || 'Not specified'),
    'UTR: ' + (data.donorUtr || 'N/A'),
    'Message: ' + (data.donorMessage || 'None'),
    'Consent for Public List: ' + (data.publicListingConsent ? 'Yes' : 'No'),
    'Payment Screenshot: ' + (paymentUrl || 'not uploaded')
  ].join('\n');

  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: 'IRCA - New Donation Received from ' + data.donorName,
      body: 'A new donation form was submitted for Rupsha.\n\n' + summary
    });
  } catch (err) {}
}

// ==========================================
// SHEET & DRIVE HELPERS
// ==========================================
function createSubmissionsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet('Submissions');
  sheet.appendRow([
    'Submitted At', 'Reference No', 'Candidate Name', 'Guardian Name', 'Address',
    'Email', 'Phone', 'WhatsApp', 'DOB', 'Gender', 'Education', 'Aadhar',
    'Profession', 'Subjects', 'Amount Payable', 'UTR', 'Photo Link', 'Payment Screenshot Link'
  ]);
  sheet.setFrozenRows(1);
  return sheet;
}

function createDonationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet('donation');
  sheet.appendRow([
    'Submitted At', 'Donor Name', 'Phone', 'Email',
    'Amount', 'UTR / Ref', 'Message', 'Public Listing Consent', 'Payment Screenshot Link'
  ]);
  sheet.setFrozenRows(1);
  return sheet;
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function saveFile(folder, fileName, base64Data, reference, kind) {
  if (!base64Data) return '';
  try {
    const contentType = guessContentType(fileName);
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, `${reference}_${kind}_${fileName}`);
    const file = folder.createFile(blob);
    // Note: Leaving file permissions to inherit from the parent folder is safer for privacy.
    return file.getUrl();
  } catch (err) {
    return 'Upload failed: ' + err.message;
  }
}

function guessContentType(fileName) {
  const ext = (fileName || '').split('.').pop().toLowerCase();
  if (ext === 'png') return 'image/png';
  if (ext === 'jpg' || ext === 'jpeg') return 'image/jpeg';
  if (ext === 'webp') return 'image/webp';
  return 'application/octet-stream';
}