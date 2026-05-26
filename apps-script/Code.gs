var SPREADSHEET_ID = '1wwIKSOVXfrtLKnLFIhNxWfL_1A2x8owVdvhpPxD8f7k';

var ROUTES = {

  'General': {
    sheet: 'info',
    to: 'info@binditobrera.it',
    from: 'info@binditobrera.it'
  },

  'Press': {
    sheet: 'info',
    to: 'info@binditobrera.it',
    from: 'info@binditobrera.it'
  },

  'Partnership': {
    sheet: 'partners',
    to: 'partners@binditobrera.it',
    from: 'partners@binditobrera.it'
  },

  'Institutional': {
    sheet: 'partners',
    to: 'partners@binditobrera.it',
    from: 'partners@binditobrera.it'
  },

  'Volunteer': {
    sheet: 'support',
    to: 'support@binditobrera.it',
    from: 'support@binditobrera.it'
  }

};

var CONTACT_HEADERS = [
  'Timestamp',
  'Name',
  'Email',
  'Enquiry Type',
  'Message'
];

var REGISTRATION_HEADERS = [
  'Timestamp',
  'Name',
  'Email',
  'Number of Attendees'
];

function doPost(e) {

  try {

    var data = JSON.parse(e.postData.contents);
    var submissionType = data.type || 'contact';

    if (submissionType === 'registration') {
      return handleRegistration(data);
    } else {
      return handleContact(data);
    }

  }

  catch (err) {

    return ContentService
      .createTextOutput(
        JSON.stringify({
          status: 'error',
          message: err.message
        })
      )
      .setMimeType(ContentService.MimeType.JSON);

  }

}

// =========================
// CONTACT FORM HANDLER
// =========================

function handleContact(data) {

  var name      = data.name || '';
  var email     = data.email || '';
  var enquiry   = data.enquiry || 'General';
  var message   = data.message || '';

  var timestamp = new Date();

  var route = ROUTES[enquiry] || ROUTES['General'];

  // =========================
  // GOOGLE SHEET STORAGE
  // =========================

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var sheet = ss.getSheetByName(route.sheet);

  if (!sheet) {
    sheet = ss.insertSheet(route.sheet);
  }

  if (sheet.getLastRow() === 0) {

    sheet.appendRow(CONTACT_HEADERS);

    sheet
      .getRange(1, 1, 1, CONTACT_HEADERS.length)
      .setFontWeight('bold');

  }

  sheet.appendRow([
    timestamp,
    name,
    email,
    enquiry,
    message
  ]);

  // =========================
  // INTERNAL TEAM EMAIL
  // =========================

  GmailApp.sendEmail(

    route.to,

    '[Bindi to Brera] ' + enquiry + ' enquiry from ' + name,

    'New enquiry received.',

    {

      from: route.from,

      name: 'Bindi to Brera',

      htmlBody: getInternalTemplate(
        name,
        email,
        enquiry,
        message,
        timestamp
      )

    }

  );

  // =========================
  // VISITOR CONFIRMATION EMAIL
  // =========================

  GmailApp.sendEmail(

    email,

    'Thank you for contacting Bindi to Brera',

    'Thank you for contacting Bindi to Brera.',

    {

      from: route.from,

      name: 'Bindi to Brera',

      htmlBody: getVisitorTemplate(
        name,
        message
      )

    }

  );

  // =========================
  // SUCCESS RESPONSE
  // =========================

  return ContentService
    .createTextOutput(
      JSON.stringify({
        status: 'success'
      })
    )
    .setMimeType(ContentService.MimeType.JSON);

}

// =========================
// REGISTRATION FORM HANDLER
// =========================

function handleRegistration(data) {

  var name      = data.name || '';
  var email     = data.email || '';
  var attendees = data.attendees || 1;

  var timestamp = new Date();

  // =========================
  // GOOGLE SHEET STORAGE
  // =========================

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var sheet = ss.getSheetByName('registrations');

  if (!sheet) {
    sheet = ss.insertSheet('registrations');
  }

  if (sheet.getLastRow() === 0) {

    sheet.appendRow(REGISTRATION_HEADERS);

    sheet
      .getRange(1, 1, 1, REGISTRATION_HEADERS.length)
      .setFontWeight('bold');

  }

  sheet.appendRow([
    timestamp,
    name,
    email,
    attendees
  ]);

  // =========================
  // INTERNAL TEAM EMAIL
  // =========================

  GmailApp.sendEmail(

    'support@binditobrera.it',

    '[Bindi to Brera] New registration from ' + name,

    'New event registration received.',

    {

      from: 'support@binditobrera.it',

      name: 'Bindi to Brera',

      htmlBody: getRegistrationInternalTemplate(
        name,
        email,
        attendees,
        timestamp
      )

    }

  );

  // =========================
  // VISITOR CONFIRMATION EMAIL
  // =========================

  GmailApp.sendEmail(

    email,

    'Registration confirmed — Bindi to Brera',

    'Thank you for registering for Bindi to Brera.',

    {

      from: 'support@binditobrera.it',

      name: 'Bindi to Brera',

      htmlBody: getRegistrationVisitorTemplate(
        name,
        attendees
      )

    }

  );

  // =========================
  // SUCCESS RESPONSE
  // =========================

  return ContentService
    .createTextOutput(
      JSON.stringify({
        status: 'success'
      })
    )
    .setMimeType(ContentService.MimeType.JSON);

}
