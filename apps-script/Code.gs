var SPREADSHEET_ID = '1wwIKSOVXfrtLKnLFIhNxWfL_1A2x8owVdvhpPxD8f7k';

var ROUTES = {
  'General':      { sheet: 'info',     to: 'info@binditobrera.it',     from: 'info@binditobrera.it'     },
  'Generale':     { sheet: 'info',     to: 'info@binditobrera.it',     from: 'info@binditobrera.it'     },
  'Press':        { sheet: 'info',     to: 'info@binditobrera.it',     from: 'info@binditobrera.it'     },
  'Stampa':       { sheet: 'info',     to: 'info@binditobrera.it',     from: 'info@binditobrera.it'     },
  'Partnership':  { sheet: 'partners', to: 'partners@binditobrera.it', from: 'partners@binditobrera.it' },
  'Institutional':{ sheet: 'partners', to: 'partners@binditobrera.it', from: 'partners@binditobrera.it' },
  'Istituzionale':{ sheet: 'partners', to: 'partners@binditobrera.it', from: 'partners@binditobrera.it' },
  'Volunteer':    { sheet: 'support',  to: 'support@binditobrera.it',  from: 'support@binditobrera.it'  },
  'Volontariato': { sheet: 'support',  to: 'support@binditobrera.it',  from: 'support@binditobrera.it'  }
};

var CONTACT_HEADERS      = ['Timestamp', 'Name', 'Email', 'Enquiry Type', 'Message'];
var REGISTRATION_HEADERS = ['Timestamp', 'Ref',  'Name', 'Email', 'Attendees'];

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

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(route.sheet) || ss.insertSheet(route.sheet);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(CONTACT_HEADERS);
    sheet.getRange(1, 1, 1, CONTACT_HEADERS.length).setFontWeight('bold');
  }
  sheet.appendRow([timestamp, name, email, enquiry, message]);

  GmailApp.sendEmail(
    route.to,
    '[Bindi to Brera] ' + enquiry + ' enquiry from ' + name,
    'New enquiry received.',
    { from: route.from, name: 'Bindi to Brera', htmlBody: contactInternalHtml(name, email, enquiry, message, timestamp) }
  );

  GmailApp.sendEmail(
    email,
    'Grazie per averci contattato — Bindi to Brera',
    'Abbiamo ricevuto il tuo messaggio e ti risponderemo al più presto.',
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

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('registrations') || ss.insertSheet('registrations');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(REGISTRATION_HEADERS);
    sheet.getRange(1, 1, 1, REGISTRATION_HEADERS.length).setFontWeight('bold');
  }
  sheet.appendRow([timestamp, ref, name, email, attendees]);

  // QR generation is best-effort — emails always send even if it fails
  var qrBlob = null;
  try { qrBlob = generateQRBlob(name, attendees, ref); } catch(e) {}

  GmailApp.sendEmail(
    'support@binditobrera.it',
    '[Bindi to Brera] Nuova registrazione — ' + name + ' · ' + ref,
    'Nuova registrazione ricevuta.',
    { from: 'support@binditobrera.it', name: 'Bindi to Brera', htmlBody: registrationInternalHtml(name, email, attendees, ref, timestamp) }
  );

  var visitorOpts = {
    from: 'support@binditobrera.it',
    name: 'Bindi to Brera',
    htmlBody: registrationVisitorHtml(name, attendees, ref, qrBlob !== null)
  };
  if (qrBlob) visitorOpts.inlineImages = { registrationQR: qrBlob };

  GmailApp.sendEmail(
    email,
    'Registrazione confermata — Bindi to Brera · ' + ref,
    'La tua registrazione per Bindi to Brera è confermata. Ref: ' + ref,
    visitorOpts
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

  // Generate at 400px for crisp retina rendering, display at 200px
  var url = 'https://api.qrserver.com/v1/create-qr-code/'
    + '?size=400x400'
    + '&data=' + encodeURIComponent(qrData)
    + '&color=171411'
    + '&bgcolor=f3ecdc'
    + '&margin=16'
    + '&qzone=1';

  return UrlFetchApp.fetch(url).getBlob().setName('bindi-to-brera-qr.png');
}

// ─────────────────────────────────────────────
// EMAIL TEMPLATES
// ─────────────────────────────────────────────
//
// Layout: fully single-column — works at any width.
// No side-by-side columns, so nothing breaks on mobile.
// ─────────────────────────────────────────────

function emailShell(previewText, bodyContent) {
  return '<!DOCTYPE html>'
    + '<html lang="it">'
    + '<head>'
    + '<meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<meta name="x-apple-disable-message-reformatting">'
    + '<title>Bindi to Brera</title>'
    + '<!--[if mso]><noscript><xml><o:OfficeDocumentSettings><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml></noscript><![endif]-->'
    + '<style>'
    + 'body,table,td,p,a{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%}'
    + 'table,td{mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse}'
    + 'img{border:0;line-height:100%;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic}'
    + 'a[x-apple-data-detectors]{color:inherit!important;text-decoration:none!important}'
    + '@media only screen and (max-width:480px){'
    + '.email-container{width:100%!important}'
    + '.body-pad{padding:28px 20px!important}'
    + '.hero-title{font-size:26px!important;line-height:1.15!important}'
    + '.info-row td{padding:10px 0!important}'
    + '.qr-img{width:180px!important;height:180px!important}'
    + '}'
    + '</style>'
    + '</head>'
    + '<body style="margin:0;padding:0;background-color:#f3ecdc;word-break:break-word">'
    // Preview text (hidden)
    + '<div style="display:none;font-size:1px;color:#f3ecdc;line-height:1px;max-height:0;max-width:0;opacity:0;overflow:hidden">' + previewText + '</div>'
    // Outer wrapper
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#f3ecdc">'
    + '<tr><td align="center" style="padding:32px 16px">'
    // Container
    + '<table role="presentation" class="email-container" cellpadding="0" cellspacing="0" style="width:100%;max-width:580px">'
    // ── Header ──
    + '<tr><td style="background-color:#0f0d0a;padding:24px 36px 20px">'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">'
    + '<tr>'
    + '<td>'
    + '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:700;letter-spacing:4px;text-transform:uppercase;color:#DAA520;line-height:1">BINDI TO BRERA</p>'
    + '<p style="margin:5px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:9px;letter-spacing:3px;text-transform:uppercase;color:rgba(243,236,220,.4);line-height:1">Unfiltered India</p>'
    + '</td>'
    + '<td align="right" style="vertical-align:middle">'
    + '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:8px;letter-spacing:2px;text-transform:uppercase;color:rgba(243,236,220,.3);line-height:1.4;text-align:right">Milano<br>11–12 Lug. 2026</p>'
    + '</td>'
    + '</tr>'
    + '</table>'
    + '</td></tr>'
    // Accent line
    + '<tr><td style="background-color:#D06224;height:3px;font-size:0;line-height:0">&nbsp;</td></tr>'
    // ── Body ──
    + '<tr><td class="body-pad" style="background-color:#ffffff;padding:40px 36px">'
    + bodyContent
    + '</td></tr>'
    // ── Footer ──
    + '<tr><td style="background-color:#0f0d0a;padding:20px 36px">'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">'
    + '<tr>'
    + '<td>'
    + '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:9px;letter-spacing:2px;text-transform:uppercase;color:rgba(243,236,220,.3)">© 2026 Bindi to Brera</p>'
    + '<p style="margin:3px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:9px;color:rgba(243,236,220,.2)">Fabbrica del Vapore, Milano, Italia</p>'
    + '</td>'
    + '<td align="right">'
    + '<a href="mailto:info@binditobrera.it" style="font-family:Arial,Helvetica,sans-serif;font-size:9px;color:rgba(210,98,36,.6);text-decoration:none">info@binditobrera.it</a>'
    + '</td>'
    + '</tr>'
    + '</table>'
    + '</td></tr>'
    + '</table>'
    + '</td></tr>'
    + '</table>'
    + '</body></html>';
}

// ── Shared building blocks ─────────────────────

function infoRow(labelText, valueText) {
  return '<tr class="info-row">'
    + '<td style="padding:11px 0;border-bottom:1px solid #f0ebe0;vertical-align:top;width:38%">'
    + '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:8px;font-weight:700;letter-spacing:2.5px;text-transform:uppercase;color:#D06224;line-height:1">' + labelText + '</p>'
    + '</td>'
    + '<td style="padding:11px 0 11px 16px;border-bottom:1px solid #f0ebe0;vertical-align:top">'
    + '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#0f0d0a;line-height:1.45;word-break:break-word">' + valueText + '</p>'
    + '</td>'
    + '</tr>';
}

function sectionLabel(text) {
  return '<p style="margin:0 0 6px;font-family:Arial,Helvetica,sans-serif;font-size:8px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:#D06224">' + text + '</p>';
}

function dashedDivider() {
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:28px 0">'
    + '<tr><td style="border-top:2px dashed rgba(15,13,10,.12);font-size:0;line-height:0">&nbsp;</td></tr>'
    + '</table>';
}

function solidDivider() {
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:28px 0">'
    + '<tr><td style="border-top:1px solid rgba(15,13,10,.1);font-size:0;line-height:0">&nbsp;</td></tr>'
    + '</table>';
}

// ── Registration: visitor confirmation ─────────

function registrationVisitorHtml(name, attendees, ref, hasQR) {
  var attendeeLabel = attendees + ' ' + (attendees == 1 ? 'persona' : 'persone');

  var qrBlock = hasQR
    // QR ticket block — when image is available
    ? '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">'
      + '<tr><td align="center" style="background-color:#f3ecdc;border:1px solid rgba(15,13,10,.1);padding:28px 20px">'
      + '<img class="qr-img" src="cid:registrationQR" width="200" height="200" alt="QR di accesso" style="display:block;margin:0 auto;border:0">'
      + '<p style="margin:14px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:8px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:rgba(15,13,10,.4)">Mostra all\'ingresso &nbsp;·&nbsp; Show at entry</p>'
      + '</td></tr>'
      + '</table>'
    // Fallback — plain ref when QR could not be generated
    : '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">'
      + '<tr><td align="center" style="background-color:#f3ecdc;border:1px solid rgba(15,13,10,.1);padding:28px 20px">'
      + '<p style="margin:0 0 6px;font-family:Arial,Helvetica,sans-serif;font-size:8px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:#D06224">Codice di accesso · Entry code</p>'
      + '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:26px;font-weight:800;color:#0f0d0a;letter-spacing:2px">' + ref + '</p>'
      + '<p style="margin:10px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:8px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:rgba(15,13,10,.4)">Mostra all\'ingresso &nbsp;·&nbsp; Show at entry</p>'
      + '</td></tr>'
      + '</table>';

  var body = ''
    // Tag + headline
    + sectionLabel('Registrazione confermata · Registration confirmed')
    + '<h1 class="hero-title" style="margin:10px 0 32px;font-family:Arial,Helvetica,sans-serif;font-size:30px;font-weight:800;color:#0f0d0a;line-height:1.1;letter-spacing:-0.5px">'
    + 'Ci vediamo<br>a Milano.</h1>'

    + qrBlock

    + dashedDivider()

    // Registration details
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">'
    + infoRow('Riferimento · Ref', '<strong style="font-family:Arial,Helvetica,sans-serif;letter-spacing:1px">' + ref + '</strong>')
    + infoRow('Nome · Name', name)
    + infoRow('Partecipanti · Attendees', attendeeLabel)
    + '</table>'

    + dashedDivider()

    // Event details
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">'
    + infoRow('Data · Date', '11–12 Luglio / July 2026')
    + infoRow('Sede · Venue', 'Fabbrica del Vapore<br><span style="color:rgba(15,13,10,.5);font-size:13px">Milano, Italia</span>')
    + infoRow('Ingresso · Admission', 'Gratuito · Free')
    + infoRow('Formato · Format', 'Mostra pubblica · Public exhibition')
    + '</table>'

    + solidDivider()

    // Instructions
    + '<p style="margin:0 0 6px;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#0f0d0a;line-height:1.7">'
    + 'Salva questo QR sul tuo telefono e mostralо all\'ingresso. Non vediamo l\'ora di accoglierti!</p>'
    + '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:rgba(15,13,10,.45);line-height:1.7">'
    + 'Save this QR on your phone and show it at the entrance. We look forward to welcoming you!</p>';

  return emailShell(
    'La tua registrazione per Bindi to Brera è confermata. Ref: ' + ref,
    body
  );
}

// ── Registration: internal team notification ───

function registrationInternalHtml(name, email, attendees, ref, timestamp) {
  var body = ''
    + sectionLabel('Nuova registrazione')
    + '<h2 style="margin:8px 0 28px;font-family:Arial,Helvetica,sans-serif;font-size:22px;font-weight:700;color:#0f0d0a">' + name + '</h2>'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">'
    + infoRow('Ref', ref)
    + infoRow('Nome', name)
    + infoRow('Email', '<a href="mailto:' + email + '" style="color:#D06224;text-decoration:none">' + email + '</a>')
    + infoRow('Partecipanti', attendees + '')
    + infoRow('Timestamp', timestamp.toLocaleString('it-IT'))
    + '</table>';

  return emailShell('Nuova registrazione: ' + name + ' · ' + ref, body);
}

// ── Contact: visitor confirmation ─────────────

function contactVisitorHtml(name) {
  var body = ''
    + sectionLabel('Messaggio ricevuto · Message received')
    + '<h1 class="hero-title" style="margin:10px 0 28px;font-family:Arial,Helvetica,sans-serif;font-size:30px;font-weight:800;color:#0f0d0a;line-height:1.1;letter-spacing:-0.5px">'
    + 'Grazie,<br>' + name + '.</h1>'

    + '<p style="margin:0 0 10px;font-family:Arial,Helvetica,sans-serif;font-size:15px;color:#0f0d0a;line-height:1.75">'
    + 'Abbiamo ricevuto il tuo messaggio e ti risponderemo al più presto.</p>'
    + '<p style="margin:0 0 28px;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:rgba(15,13,10,.45);line-height:1.75">'
    + 'We have received your message and will get back to you shortly.</p>'

    + solidDivider()

    // Event reminder block
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#f3ecdc;border-left:3px solid #D06224">'
    + '<tr><td style="padding:16px 20px">'
    + '<p style="margin:0 0 3px;font-family:Arial,Helvetica,sans-serif;font-size:8px;font-weight:700;letter-spacing:2.5px;text-transform:uppercase;color:#D06224">L\'evento</p>'
    + '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#0f0d0a;line-height:1.6">'
    + 'Bindi to Brera &nbsp;·&nbsp; 11–12 Luglio 2026<br>'
    + '<span style="color:rgba(15,13,10,.5);font-size:13px">Fabbrica del Vapore, Milano</span></p>'
    + '</td></tr>'
    + '</table>';

  return emailShell('Abbiamo ricevuto il tuo messaggio. Ti risponderemo presto.', body);
}

// ── Contact: internal team notification ───────

function contactInternalHtml(name, email, enquiry, message, timestamp) {
  var body = ''
    + sectionLabel('Nuova richiesta di contatto')
    + '<h2 style="margin:8px 0 28px;font-family:Arial,Helvetica,sans-serif;font-size:22px;font-weight:700;color:#0f0d0a">' + name + '</h2>'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">'
    + infoRow('Nome', name)
    + infoRow('Email', '<a href="mailto:' + email + '" style="color:#D06224;text-decoration:none">' + email + '</a>')
    + infoRow('Tipo richiesta', enquiry)
    + infoRow('Timestamp', timestamp.toLocaleString('it-IT'))
    + '</table>'
    + solidDivider()
    + sectionLabel('Messaggio')
    + '<p style="margin:8px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#0f0d0a;line-height:1.75;word-break:break-word">'
    + message.replace(/\n/g, '<br>')
    + '</p>';

  return emailShell('Nuova richiesta da ' + name + ' — ' + enquiry, body);
}
