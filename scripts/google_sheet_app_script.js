/**
 * IRCA Student Admission Form — Google Apps Script backend
 * ----------------------------------------------------------
 * This script receives form submissions from student-registration.html
 * and writes them as new rows to a Google Sheet. It also saves the
 * uploaded photo and payment screenshot to a Google Drive folder and
 * stores links to those files in the sheet.
 *
 * SETUP: see the steps shared alongside this file.
 */

// Name of the folder (in your Drive) where uploaded files will be stored.
const DRIVE_FOLDER_NAME = 'IRCA Admission Uploads';

// ---- Email notification settings ----
// Email(s) that should get notified of every new admission (comma-separate for multiple).
const ADMIN_EMAIL = 'irca.admin@gmail.com,riyadeb.work@gmail.com,dipsraj.kundu@gmail.com';
// Set to false if you don't want an admin notification email at all.
const SEND_ADMIN_EMAIL = true;
// Set to true to also send a confirmation email to the applicant (uses the "Mail ID" field they entered).
const SEND_APPLICANT_CONFIRMATION = true;

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Submissions')
      || createSubmissionsSheet();

    const folder = getOrCreateFolder(DRIVE_FOLDER_NAME);

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

    sendNotificationEmails(data, photoUrl, paymentUrl);

    return ContentService.createTextOutput(JSON.stringify({ result: 'success', referenceNo: data.referenceNo }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sendNotificationEmails(data, photoUrl, paymentUrl) {
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
    } catch (err) {
      // Don't let an email failure block the submission itself.
    }
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
    } catch (err) {
      // Don't let an email failure block the submission itself.
    }
  }
}

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

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function saveFile(folder, fileName, base64Data, referenceNo, kind) {
  if (!base64Data) return '';
  try {
    const contentType = guessContentType(fileName);
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, `${referenceNo}_${kind}_${fileName}`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
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