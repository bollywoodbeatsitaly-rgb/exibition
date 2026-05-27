var SPREADSHEET_ID = '1wwIKSOVXfrtLKnLFIhNxWfL_1A2x8owVdvhpPxD8f7k';

var ROUTES = {
  'General':     { sheet: 'info',     to: 'info@binditobrera.it',     from: 'info@binditobrera.it'     },
  'Generale':    { sheet: 'info',     to: 'info@binditobrera.it',     from: 'info@binditobrera.it'     },
  'Press':       { sheet: 'info',     to: 'info@binditobrera.it',     from: 'info@binditobrera.it'     },
  'Stampa':      { sheet: 'info',     to: 'info@binditobrera.it',     from: 'info@binditobrera.it'     },
  'Partnership': { sheet: 'partners', to: 'partners@binditobrera.it', from: 'partners@binditobrera.it' },
  'Institutional':{ sheet: 'partners',to: 'partners@binditobrera.it', from: 'partners@binditobrera.it' },
  'Istituzionale':{ sheet: 'partners',to: 'partners@binditobrera.it', from: 'partners@binditobrera.it' },
  'Volunteer':   { sheet: 'support',  to: 'support@binditobrera.it',  from: 'support@binditobrera.it'  },
  'Volontariato':{ sheet: 'support',  to: 'support@binditobrera.it',  from: 'support@binditobrera.it'  }
};

var CONTACT_HEADERS      = ['Timestamp','Name','Email','Enquiry Type','Message'];
var REGISTRATION_HEADERS = ['Timestamp','Ref','Name','Email','Attendees'];

// ─────────────────────────────────────────────
// ENTRY POINT
// ─────────────────────────────────────────────

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    return data.type === 'registration' ? handleRegistration(data) : handleContact(data);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────────
// CONTACT FORM
// ─────────────────────────────────────────────

function handleContact(data) {
  var name      = data.name    || '';
  var email     = data.email   || '';
  var enquiry   = data.enquiry || 'General';
  var message   = data.message || '';
  var timestamp = new Date();
  var route     = ROUTES[enquiry] || ROUTES['General'];

  // Sheet
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(route.sheet) || ss.insertSheet(route.sheet);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(CONTACT_HEADERS);
    sheet.getRange(1, 1, 1, CONTACT_HEADERS.length).setFontWeight('bold');
  }
  sheet.appendRow([timestamp, name, email, enquiry, message]);

  // Internal email
  GmailApp.sendEmail(
    route.to,
    '[Bindi to Brera] ' + enquiry + ' enquiry from ' + name,
    'New enquiry received.',
    { from: route.from, name: 'Bindi to Brera', htmlBody: contactInternalHtml(name, email, enquiry, message, timestamp) }
  );

  // Visitor confirmation
  GmailApp.sendEmail(
    email,
    'Grazie per averci contattato — Bindi to Brera',
    'Thank you for contacting Bindi to Brera.',
    { from: route.from, name: 'Bindi to Brera', htmlBody: contactVisitorHtml(name) }
  );

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────
// REGISTRATION FORM
// ─────────────────────────────────────────────

function handleRegistration(data) {
  var name      = data.name      || '';
  var email     = data.email     || '';
  var attendees = data.attendees || 1;
  var timestamp = new Date();
  var ref       = 'BTB-' + timestamp.getTime().toString().slice(-6);

  // Sheet
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('registrations') || ss.insertSheet('registrations');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(REGISTRATION_HEADERS);
    sheet.getRange(1, 1, 1, REGISTRATION_HEADERS.length).setFontWeight('bold');
  }
  sheet.appendRow([timestamp, ref, name, email, attendees]);

  // Generate QR code blob
  var qrBlob = generateQRBlob(name, attendees, ref);

  // Internal email
  GmailApp.sendEmail(
    'support@binditobrera.it',
    '[Bindi to Brera] Nuova registrazione — ' + name + ' (' + ref + ')',
    'New registration received.',
    { from: 'support@binditobrera.it', name: 'Bindi to Brera', htmlBody: registrationInternalHtml(name, email, attendees, ref, timestamp) }
  );

  // Visitor confirmation with QR
  GmailApp.sendEmail(
    email,
    'Registrazione confermata — Bindi to Brera (' + ref + ')',
    'La tua registrazione per Bindi to Brera è confermata. / Your registration for Bindi to Brera is confirmed.',
    {
      from: 'support@binditobrera.it',
      name: 'Bindi to Brera',
      htmlBody: registrationVisitorHtml(name, attendees, ref),
      inlineImages: { registrationQR: qrBlob }
    }
  );

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────
// QR CODE GENERATOR
// ─────────────────────────────────────────────

function generateQRBlob(name, attendees, ref) {
  var qrData = [
    'BINDI TO BRERA',
    'Ref: ' + ref,
    'Name: ' + name,
    'Attendees: ' + attendees,
    'Date: 11-12 July 2026',
    'Venue: Fabbrica del Vapore, Milan'
  ].join('\n');

  var url = 'https://api.qrserver.com/v1/create-qr-code/'
    + '?size=280x280'
    + '&data=' + encodeURIComponent(qrData)
    + '&color=171411'
    + '&bgcolor=f3ecdc'
    + '&margin=12'
    + '&qzone=1';

  return UrlFetchApp.fetch(url).getBlob().setName('bindi-to-brera-qr.png');
}

// ─────────────────────────────────────────────
// EMAIL TEMPLATES
// ─────────────────────────────────────────────

function emailWrap(bodyContent) {
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#f3ecdc;font-family:Arial,sans-serif">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#f3ecdc;padding:40px 0">'
    + '<tr><td align="center">'
    + '<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%">'
    // Header
    + '<tr><td style="background:#0f0d0a;padding:28px 40px">'
    + '<p style="margin:0;font-family:Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:4px;text-transform:uppercase;color:#DAA520">BINDI TO BRERA</p>'
    + '<p style="margin:4px 0 0;font-family:Arial,sans-serif;font-size:9px;letter-spacing:3px;text-transform:uppercase;color:rgba(243,236,220,.45)">Milano · 11-12 Luglio 2026</p>'
    + '</td></tr>'
    // Body
    + '<tr><td style="background:#ffffff;padding:40px">'
    + bodyContent
    + '</td></tr>'
    // Footer
    + '<tr><td style="background:#0f0d0a;padding:20px 40px;text-align:center">'
    + '<p style="margin:0;font-family:Arial,sans-serif;font-size:9px;letter-spacing:2px;text-transform:uppercase;color:rgba(243,236,220,.35)">© 2026 Bindi to Brera · Fabbrica del Vapore, Milano</p>'
    + '</td></tr>'
    + '</table>'
    + '</td></tr></table>'
    + '</body></html>';
}

function label(text) {
  return '<p style="margin:0;font-family:Arial,sans-serif;font-size:9px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:#D06224">' + text + '</p>';
}

function value(text) {
  return '<p style="margin:4px 0 20px;font-family:Arial,sans-serif;font-size:15px;color:#0f0d0a">' + text + '</p>';
}

function divider() {
  return '<hr style="border:none;border-top:1px solid rgba(15,13,10,.1);margin:24px 0">';
}

// ── Registration: visitor confirmation ────────

function registrationVisitorHtml(name, attendees, ref) {
  var body =
    '<p style="margin:0 0 4px;font-family:Arial,sans-serif;font-size:9px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:#D06224">REGISTRAZIONE CONFERMATA · REGISTRATION CONFIRMED</p>'
    + '<h1 style="margin:8px 0 24px;font-family:Arial,sans-serif;font-size:28px;font-weight:800;color:#0f0d0a;line-height:1.1">Ci vediamo<br>a Milano.</h1>'

    // QR section
    + '<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:28px">'
    + '<tr>'
    + '<td style="background:#f3ecdc;border:1px solid rgba(15,13,10,.1);padding:24px;text-align:center;width:200px;vertical-align:middle">'
    + '<img src="cid:registrationQR" width="200" height="200" style="display:block;margin:0 auto" alt="QR di accesso">'
    + '<p style="margin:10px 0 0;font-family:Arial,sans-serif;font-size:8px;letter-spacing:2px;text-transform:uppercase;color:rgba(15,13,10,.45)">MOSTRA ALL\'INGRESSO · SHOW AT ENTRY</p>'
    + '</td>'
    + '<td style="vertical-align:top;padding-left:24px">'
    + label('Ref. registrazione') + value(ref)
    + label('Nome · Name') + value(name)
    + label('Partecipanti · Attendees') + value(attendees + (attendees === 1 ? ' persona' : ' persone'))
    + '</td>'
    + '</tr>'
    + '</table>'

    + divider()

    // Event details
    + '<table width="100%" cellpadding="0" cellspacing="0">'
    + '<tr>'
    + '<td style="width:50%;vertical-align:top">'
    + label('Data · Date') + value('11-12 Luglio / July 2026')
    + label('Sede · Venue') + value('Fabbrica del Vapore<br>Milano, Italia')
    + '</td>'
    + '<td style="width:50%;vertical-align:top;padding-left:16px">'
    + label('Ingresso · Admission') + value('Gratuito · Free')
    + label('Formato') + value('Mostra pubblica · Public exhibition')
    + '</td>'
    + '</tr>'
    + '</table>'

    + divider()

    + '<p style="margin:0 0 8px;font-family:Arial,sans-serif;font-size:13px;color:#0f0d0a">Salva questo QR sul tuo telefono e mostralo all\'ingresso. Buona visita!</p>'
    + '<p style="margin:0;font-family:Arial,sans-serif;font-size:12px;color:rgba(15,13,10,.5)">Save this QR on your phone and show it at the entrance. Enjoy the exhibition!</p>'

    + divider()

    + '<p style="margin:0;font-family:Arial,sans-serif;font-size:11px;color:rgba(15,13,10,.4)">Per informazioni: <a href="mailto:info@binditobrera.it" style="color:#D06224">info@binditobrera.it</a></p>';

  return emailWrap(body);
}

// ── Registration: internal team notification ───

function registrationInternalHtml(name, email, attendees, ref, timestamp) {
  var body =
    '<p style="margin:0 0 16px;font-family:Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:#D06224">NUOVA REGISTRAZIONE</p>'
    + label('Ref') + value(ref)
    + label('Name') + value(name)
    + label('Email') + value('<a href="mailto:' + email + '" style="color:#D06224">' + email + '</a>')
    + label('Attendees') + value(attendees)
    + label('Timestamp') + value(timestamp.toLocaleString('it-IT'));

  return emailWrap(body);
}

// ── Contact: visitor confirmation ─────────────

function contactVisitorHtml(name) {
  var body =
    '<p style="margin:0 0 4px;font-family:Arial,sans-serif;font-size:9px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:#D06224">MESSAGGIO RICEVUTO</p>'
    + '<h1 style="margin:8px 0 24px;font-family:Arial,sans-serif;font-size:28px;font-weight:800;color:#0f0d0a;line-height:1.1">Grazie,<br>' + name + '.</h1>'
    + '<p style="margin:0 0 12px;font-family:Arial,sans-serif;font-size:14px;color:#0f0d0a;line-height:1.7">Abbiamo ricevuto il tuo messaggio e ti risponderemo al più presto.</p>'
    + '<p style="margin:0 0 28px;font-family:Arial,sans-serif;font-size:13px;color:rgba(15,13,10,.5);line-height:1.7">We have received your message and will get back to you shortly.</p>'
    + divider()
    + '<p style="margin:0;font-family:Arial,sans-serif;font-size:13px;color:rgba(15,13,10,.55)">Bindi to Brera · 11-12 Luglio 2026 · Fabbrica del Vapore, Milano</p>';

  return emailWrap(body);
}

// ── Contact: internal team notification ───────

function contactInternalHtml(name, email, enquiry, message, timestamp) {
  var body =
    '<p style="margin:0 0 16px;font-family:Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:#D06224">NUOVA RICHIESTA DI CONTATTO</p>'
    + label('Name') + value(name)
    + label('Email') + value('<a href="mailto:' + email + '" style="color:#D06224">' + email + '</a>')
    + label('Enquiry Type') + value(enquiry)
    + label('Message') + value(message)
    + label('Timestamp') + value(timestamp.toLocaleString('it-IT'));

  return emailWrap(body);
}
