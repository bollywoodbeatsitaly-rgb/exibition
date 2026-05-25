var SPREADSHEET_ID = '1wwIKSOVXfrtLKnLFIhNxWfL_1A2x8owVdvhpPxD8f7k';

var ROUTES = {
  'General':       { sheet: 'info',     to: 'info@binditobrera.it' },
  'Press':         { sheet: 'info',     to: 'info@binditobrera.it' },
  'Partnership':   { sheet: 'partners', to: 'partners@binditobrera.it' },
  'Institutional': { sheet: 'partners', to: 'partners@binditobrera.it' },
  'Volunteer':     { sheet: 'support',  to: 'support@binditobrera.it' }
};

var HEADERS = ['Timestamp', 'Name', 'Email', 'Enquiry Type', 'Message'];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
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
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
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

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Keep the Web App alive and testable via GET
function doGet() {
  return ContentService
    .createTextOutput('Bindi to Brera contact endpoint is live.')
    .setMimeType(ContentService.MimeType.TEXT);
}
