var SPREADSHEET_ID = '1wwIKSOVXfrtLKnLFIhNxWfL_1A2x8owVdvhpPxD8f7k';

var ROUTES = {
  'General':       { sheet: 'info',     to: 'info@binditobrera.it' },
  'Press':         { sheet: 'info',     to: 'info@binditobrera.it' },
  'Partnership':   { sheet: 'partners', to: 'partners@binditobrera.it' },
  'Institutional': { sheet: 'partners', to: 'partners@binditobrera.it' },
  'Volunteer':     { sheet: 'support',  to: 'support@binditobrera.it' }
};

var CONTACT_HEADERS = ['Timestamp', 'Name', 'Email', 'Enquiry Type', 'Message'];
var REGISTRATION_HEADERS = ['Timestamp', 'Name', 'Email', 'Number of Attendees'];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var submissionType = data.type || 'contact';

    if (submissionType === 'registration') {
      return handleRegistration(data);
    } else {
      return handleContact(data);
    }
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleContact(data) {
  var name    = data.name    || '';
  var email   = data.email   || '';
  var enquiry = data.enquiry || 'General';
  var message = data.message || '';
  var timestamp = new Date();

  var route = ROUTES[enquiry] || ROUTES['General'];

  // Write to the correct sheet tab
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(route.sheet);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(CONTACT_HEADERS);
    sheet.getRange(1, 1, 1, CONTACT_HEADERS.length).setFontWeight('bold');
  }
  sheet.appendRow([timestamp, name, email, enquiry, message]);

  // Notify the right team inbox
  MailApp.sendEmail({
    to: route.to,
    subject: '[Bindi to Brera] ' + enquiry + ' enquiry from ' + name,
    body: [
      'New contact form submission',
      '',
      'Name:    ' + name,
      'Email:   ' + email,
      'Type:    ' + enquiry,
      'Date:    ' + timestamp.toLocaleString('en-GB'),
      '',
      'Message:',
      message,
      '',
      '— Bindi to Brera automated notification'
    ].join('\n')
  });

  // Confirmation to the visitor
  MailApp.sendEmail({
    to: email,
    subject: 'Thank you for reaching out — Bindi to Brera',
    body: [
      'Dear ' + name + ',',
      '',
      'Thank you for contacting Bindi to Brera. We have received your message and will be in touch shortly.',
      '',
      'Your message:',
      '────────────────────────────',
      message,
      '────────────────────────────',
      '',
      'Bindi to Brera',
      'Fabbrica del Vapore, Milan  |  11–12 July 2026',
      'info@binditobrera.it'
    ].join('\n')
  });

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleRegistration(data) {
  var name    = data.name    || '';
  var email   = data.email   || '';
  var attendees = data.attendees || 1;
  var timestamp = new Date();

  // Write to registrations sheet tab
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('registrations');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(REGISTRATION_HEADERS);
    sheet.getRange(1, 1, 1, REGISTRATION_HEADERS.length).setFontWeight('bold');
  }
  sheet.appendRow([timestamp, name, email, attendees]);

  // Notify support team
  MailApp.sendEmail({
    to: 'support@binditobrera.it',
    subject: '[Bindi to Brera] New registration from ' + name,
    body: [
      'New event registration received',
      '',
      'Name:       ' + name,
      'Email:      ' + email,
      'Attendees:  ' + attendees,
      'Date:       ' + timestamp.toLocaleString('en-GB'),
      '',
      '— Bindi to Brera automated notification'
    ].join('\n')
  });

  // Confirmation to the visitor
  MailApp.sendEmail({
    to: email,
    subject: 'Registration confirmed — Bindi to Brera',
    body: [
      'Dear ' + name + ',',
      '',
      'Thank you for registering for Bindi to Brera! We look forward to welcoming you and ' + (attendees - 1) + ' guest' + (attendees === 1 ? '' : 's') + '.',
      '',
      'Registration Details:',
      '────────────────────────────',
      'Name:       ' + name,
      'Attendees:  ' + attendees,
      '────────────────────────────',
      '',
      'Exhibition Details:',
      'Date: 11–12 July 2026',
      'Venue: Fabbrica del Vapore, Milan',
      'Admission: Free (no tickets required)',
      '',
      'See you soon!',
      '',
      'Bindi to Brera',
      'info@binditobrera.it'
    ].join('\n')
  });

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Keep the Web App alive and testable via GET
function doGet() {
  return ContentService
    .createTextOutput('Bindi to Brera contact endpoint is live.')
    .setMimeType(ContentService.MimeType.TEXT);
}
