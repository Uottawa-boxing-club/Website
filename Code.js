/***** CONFIG *****/
var SHEET_ID = '1HyoKB0AxA3plMrFhzcYze7JmhrHZEawbjvwND-vUaFM';

// Auto-detect these registration tab names, in order:
var REG_TAB_CANDIDATES = ['Registrations','Sheet1','Registration','Signups'];
// Waitlist tab name:
var WAIT_TAB_NAME = 'Waitlist';
// Dashboard tab name
var DASHBOARD_TAB_NAME = 'Dashboard';

// Seats per class (NOW per SESSION)
var CAPACITY_PER_CLASS = 22;

// Master email log tab (unique emails for the semester)
var MASTER_EMAIL_TAB_NAME = 'Semester Emails';

/***** DATE-BASED SESSION SCHEDULE (WINTER 2026) *****
 * IMPORTANT: each session is its own "classChoice" key,
 * so capacity/waitlist is per exact date+time.
 */
var DATE_BASED_SCHEDULE = [
  // Primary – Thursday 4–5
  { iso: "2026-01-22", dayEN: "Thursday", dayFR: "Jeudi",    time: "4:00–5:00 PM",  typeEN: "Primary",  typeFR: "Principal" },
  { iso: "2026-01-29", dayEN: "Thursday", dayFR: "Jeudi",    time: "4:00–5:00 PM",  typeEN: "Primary",  typeFR: "Principal" },
  { iso: "2026-02-05", dayEN: "Thursday", dayFR: "Jeudi",    time: "4:00–5:00 PM",  typeEN: "Primary",  typeFR: "Principal" },
  { iso: "2026-02-12", dayEN: "Thursday", dayFR: "Jeudi",    time: "4:00–5:00 PM",  typeEN: "Primary",  typeFR: "Principal" },
  { iso: "2026-02-26", dayEN: "Thursday", dayFR: "Jeudi",    time: "4:00–5:00 PM",  typeEN: "Primary",  typeFR: "Principal" },
  { iso: "2026-03-05", dayEN: "Thursday", dayFR: "Jeudi",    time: "4:00–5:00 PM",  typeEN: "Primary",  typeFR: "Principal" },
  { iso: "2026-03-12", dayEN: "Thursday", dayFR: "Jeudi",    time: "4:00–5:00 PM",  typeEN: "Primary",  typeFR: "Principal" },
  { iso: "2026-03-19", dayEN: "Thursday", dayFR: "Jeudi",    time: "4:00–5:00 PM",  typeEN: "Primary",  typeFR: "Principal" },
  { iso: "2026-03-26", dayEN: "Thursday", dayFR: "Jeudi",    time: "4:00–5:00 PM",  typeEN: "Primary",  typeFR: "Principal" },
  { iso: "2026-04-02", dayEN: "Thursday", dayFR: "Jeudi",    time: "4:00–5:00 PM",  typeEN: "Primary",  typeFR: "Principal" },
  { iso: "2026-04-09", dayEN: "Thursday", dayFR: "Jeudi",    time: "4:00–5:00 PM",  typeEN: "Primary",  typeFR: "Principal" },

  // Secondary – Tuesday 4–5
  { iso: "2026-01-20", dayEN: "Tuesday",  dayFR: "Mardi",    time: "4:00–5:00 PM",  typeEN: "Secondary", typeFR: "Secondaire" },
  { iso: "2026-02-03", dayEN: "Tuesday",  dayFR: "Mardi",    time: "4:00–5:00 PM",  typeEN: "Secondary", typeFR: "Secondaire" },
  { iso: "2026-03-03", dayEN: "Tuesday",  dayFR: "Mardi",    time: "4:00–5:00 PM",  typeEN: "Secondary", typeFR: "Secondaire" },
  { iso: "2026-03-17", dayEN: "Tuesday",  dayFR: "Mardi",    time: "4:00–5:00 PM",  typeEN: "Secondary", typeFR: "Secondaire" },
  { iso: "2026-03-31", dayEN: "Tuesday",  dayFR: "Mardi",    time: "4:00–5:00 PM",  typeEN: "Secondary", typeFR: "Secondaire" },

  // Secondary – Friday 12–1
  { iso: "2026-01-30", dayEN: "Friday",   dayFR: "Vendredi", time: "12:00–1:00 PM", typeEN: "Secondary", typeFR: "Secondaire" },
  { iso: "2026-02-13", dayEN: "Friday",   dayFR: "Vendredi", time: "12:00–1:00 PM", typeEN: "Secondary", typeFR: "Secondaire" },
  { iso: "2026-02-27", dayEN: "Friday",   dayFR: "Vendredi", time: "12:00–1:00 PM", typeEN: "Secondary", typeFR: "Secondaire" },
  { iso: "2026-03-13", dayEN: "Friday",   dayFR: "Vendredi", time: "12:00–1:00 PM", typeEN: "Secondary", typeFR: "Secondaire" },
  { iso: "2026-03-27", dayEN: "Friday",   dayFR: "Vendredi", time: "12:00–1:00 PM", typeEN: "Secondary", typeFR: "Secondaire" },
  { iso: "2026-04-10", dayEN: "Friday",   dayFR: "Vendredi", time: "12:00–1:00 PM", typeEN: "Secondary", typeFR: "Secondaire" }
];

/***** SESSION KEY + DISPLAY *****/
function _sessionKey_(s) {
  // Stable key used everywhere in sheets
  return s.iso + ' | ' + s.dayEN + ' | ' + s.time;
}

// Nicely formatted (and trims spaces so it doesn't fail)
function _displaySession_(classChoice) {
  classChoice = _s(classChoice);

  // Split and trim each piece
  var parts = classChoice.split('|').map(function(p){ return _s(p); });

  if (parts.length >= 3 && /^\d{4}-\d{2}-\d{2}$/.test(parts[0])) {
    var iso = parts[0];
    var day = parts[1];
    var time = parts[2];
    return day + ' ' + iso + ' — ' + time;
  }

  return classChoice;
}

/***** MASTER EMAIL LOG (NO DUPLICATES) *****/
function _getMasterEmailSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(MASTER_EMAIL_TAB_NAME);
  if (!sh) {
    sh = ss.insertSheet(MASTER_EMAIL_TAB_NAME);
    sh.appendRow(['First Seen', 'Email', 'Name (latest)', 'Last Seen', 'Signup Count']);
  } else {
    _ensureHeaders(sh, ['First Seen', 'Email', 'Name (latest)', 'Last Seen', 'Signup Count']);
  }
  return sh;
}

function _recordSemesterEmail_(email, name) {
  email = _s(email).toLowerCase();
  name = _s(name);
  if (!email) return;

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var sh = _getMasterEmailSheet_();
    var lastRow = sh.getLastRow();

    if (lastRow < 2) {
      sh.appendRow([new Date(), email, name, new Date(), 1]);
      return;
    }

    var emails = sh.getRange(2, 2, lastRow - 1, 1).getValues(); // B2:B
    for (var i = 0; i < emails.length; i++) {
      var existing = _s(emails[i][0]).toLowerCase();
      if (existing === email) {
        var row = i + 2;
        sh.getRange(row, 3).setValue(name || sh.getRange(row, 3).getValue()); // Name (latest)
        sh.getRange(row, 4).setValue(new Date()); // Last Seen

        var countCell = sh.getRange(row, 5);
        var cur = Number(countCell.getValue() || 0);
        countCell.setValue(cur + 1);
        return;
      }
    }

    sh.appendRow([new Date(), email, name, new Date(), 1]);

  } finally {
    lock.releaseLock();
  }
}

/***** WEB APP *****/
function doGet() {
  var html;
  try {
    html = HtmlService.createHtmlOutputFromFile('Index');
  } catch (e1) {
    try {
      html = HtmlService.createHtmlOutputFromFile('index');
    } catch (e2) {
      var msg = 'HTML file not found. Create a file named "Index" (or "index") in this project.';
      return HtmlService.createHtmlOutput('<pre style="font:14px/1.4 monospace;color:#b00;">' + msg + '</pre>');
    }
  }
  return html
    .setTitle('uOttawa Boxing Club')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/***** API: schedule for frontend *****/
function getSchedule() {
  return DATE_BASED_SCHEDULE.map(function(s){
    return {
      isoDate: s.iso,
      dayEN: s.dayEN,
      dayFR: s.dayFR,
      time: s.time,
      typeEN: s.typeEN,
      typeFR: s.typeFR,
      value: _sessionKey_(s)
    };
  });
}

/***** DASHBOARD GROUPING (3 streams now) *****/
function _currentSchedule_() {
  return {
    day1: { name: 'TUESDAY',   headerReg: /tuesday/i,  sectionLabel: 'TUESDAY 4:00-5:00 PM (Secondary)' },
    day2: { name: 'THURSDAY',  headerReg: /thursday/i, sectionLabel: 'THURSDAY 4:00-5:00 PM (Primary)' },
    day3: { name: 'FRIDAY',    headerReg: /friday/i,   sectionLabel: 'FRIDAY 12:00-1:00 PM (Secondary)' }
  };
}

/***** MAIN ACTIONS *****/
function registerNew(form) {
  var name = _s(form && form.name);
  var email = _s(form && form.email).toLowerCase();
  var phone = _s(form && form.phone);
  var classChoice = _s(form && form.classChoice);
  var classChoiceDisplay = _displaySession_(classChoice);

  if (!name || !email || !phone || !classChoice) {
    return { success: false, error: 'Please provide name, email, phone, and class session.' };
  }

  var regSheet  = _getRegistrationSheet();
  var waitSheet = _getWaitlistSheet();

  var regData = regSheet.getDataRange().getValues();
  var rmap = _headerMap(regData[0] || []);
  var cEmail = _idx(rmap, ['email','e-mail']);
  var cClass = _idx(rmap, ['class choice','class','classchoice']);
  if (cEmail < 0 || cClass < 0) {
    return { success:false, error:'Header mismatch in Registrations. Expect "Email" & "Class Choice".' };
  }

  // Prevent duplicate (same email + session)
  for (var i = 1; i < regData.length; i++) {
    var e = _s(regData[i][cEmail]).toLowerCase();
    var c = _s(regData[i][cClass]);
    if (e === email && c === classChoice) {
      try { _sendEmailRegistered(email, name, classChoiceDisplay); } catch (e1) {}
      _safeRebuildDashboard();
      return { success:true, status:'registered', message:'You are already registered for this class.' };
    }
  }

  // Count active in this session
  var active = 0;
  for (var j = 1; j < regData.length; j++) {
    var c2 = _s(regData[j][cClass]);
    if (c2 === classChoice) active++;
  }

  // Registered
  if (active < CAPACITY_PER_CLASS) {
    regSheet.appendRow([new Date(), name, email, phone, classChoice]);
    _recordSemesterEmail_(email, name);
    try { _sendEmailRegistered(email, name, classChoiceDisplay); } catch (e2) {}
    _safeRebuildDashboard();
    return { success:true, status:'registered', message:'Registered successfully.' };
  }

  // Waitlist position within this session
  var waitData = waitSheet.getDataRange().getValues();
  var wmap = _headerMap(waitData[0] || []);
  var wClass = _idx(wmap, ['class choice','class','classchoice']);
  var position = 1;
  for (var k = 1; k < waitData.length; k++) {
    var wc = _s(waitData[k][wClass]);
    if (wc === classChoice) position++;
  }

  waitSheet.appendRow([new Date(), name, email, phone, classChoice, position]);
  _recordSemesterEmail_(email, name);
  try { _sendEmailWaitlisted(email, name, classChoiceDisplay, position); } catch (e3) {}
  _safeRebuildDashboard();
  return {
    success:true, status:'waitlisted', position:position,
    message:'Class is full. You are on the waitlist (position ' + position + ').'
  };
}

function cancelRegistration(form) {
  var email = _s(form && form.email).toLowerCase();
  var classChoice = _s(form && form.classChoice); // optional (blank = all)

  if (!email) return { success:false, error:'Email is required to cancel.' };

  var regSheet = _getRegistrationSheet();
  var vals = regSheet.getDataRange().getValues();

  if (!vals || vals.length < 2) {
    var resWLOnly = _cancelFromWaitlistOnly_(email, classChoice);
    _safeRebuildDashboard();
    return resWLOnly;
  }

  var hmap = _headerMap(vals[0] || []);
  var cEmail = _idx(hmap, ['email','e-mail']);
  var cClass = _idx(hmap, ['class choice','class','classchoice']);
  if (cEmail < 0 || cClass < 0) {
    return { success:false, error:'Header mismatch in Registrations. Expect "Email" & "Class Choice".' };
  }

  var rowsToDelete = [];
  var freedByClass = {};
  for (var r = 1; r < vals.length; r++) {
    var e = _s(vals[r][cEmail]).toLowerCase();
    var c = _s(vals[r][cClass]);
    var match = (e === email) && (classChoice ? (c === classChoice) : true);
    if (match) {
      rowsToDelete.push({ row:r+1, cls:c });
      freedByClass[c] = (freedByClass[c] || 0) + 1;
    }
  }

  if (rowsToDelete.length === 0) {
    var resWL = _cancelFromWaitlistOnly_(email, classChoice);
    _safeRebuildDashboard();
    return resWL;
  }

  // Delete bottom-up
  for (var i = rowsToDelete.length - 1; i >= 0; i--) {
    regSheet.deleteRow(rowsToDelete[i].row);
  }

  // Promote from waitlist for each freed seat
  var promotedTotal = 0;
  for (var cls in freedByClass) {
    if (freedByClass.hasOwnProperty(cls)) {
      var seats = freedByClass[cls];
      for (var s = 0; s < seats; s++) {
        promotedTotal += _promoteFromWaitlist(cls);
      }
    }
  }

  try { _sendEmailCancelled(email, classChoice || null, rowsToDelete.length); } catch (e4) {}

  var msg = rowsToDelete.length + ' registration' + (rowsToDelete.length === 1 ? '' : 's') + ' cancelled';
  if (promotedTotal > 0) msg += '. ' + promotedTotal + ' waitlisted member' + (promotedTotal === 1 ? '' : 's') + ' promoted.';
  _safeRebuildDashboard();
  return { success:true, deleted:rowsToDelete.length, promoted:promotedTotal, message: msg };
}

/***** PROMOTION *****/
function _promoteFromWaitlist(classChoice) {
  var regSheet  = _getRegistrationSheet();
  var waitSheet = _getWaitlistSheet();

  var wData = waitSheet.getDataRange().getValues();
  if (wData.length < 2) return 0;

  var wmap = _headerMap(wData[0] || []);
  var cName  = _idx(wmap, ['name']);
  var cEmail = _idx(wmap, ['email','e-mail']);
  var cPhone = _idx(wmap, ['phone']);
  var cClass = _idx(wmap, ['class choice','class','classchoice']);

  var rowIndex = -1, rowData = null;
  for (var r = 1; r < wData.length; r++) {
    var cls = _s(wData[r][cClass]);
    if (cls === classChoice) { rowIndex = r + 1; rowData = wData[r]; break; }
  }
  if (rowIndex === -1) return 0;

  var name  = rowData[cName];
  var email = _s(rowData[cEmail]).toLowerCase();
  var phone = rowData[cPhone];

  regSheet.appendRow([new Date(), name, email, phone, classChoice]);
  waitSheet.deleteRow(rowIndex);
  _renumberWaitlist(waitSheet, classChoice);

  try { _sendEmailPromoted(email, name, _displaySession_(classChoice)); } catch (e5) {}
  return 1;
}

function _renumberWaitlist(waitSheet, classChoice) {
  var vals = waitSheet.getDataRange().getValues();
  if (vals.length < 2) return;
  var wmap = _headerMap(vals[0] || []);
  var cClass = _idx(wmap, ['class choice','class','classchoice']);
  var cPos   = _idx(wmap, ['position']);
  if (cClass < 0 || cPos < 0) return;

  var pos = 1;
  for (var r = 1; r < vals.length; r++) {
    var cls = _s(vals[r][cClass]);
    if (cls === classChoice) {
      waitSheet.getRange(r + 1, cPos + 1).setValue(pos++);
    }
  }
}

/***** WAITLIST CANCEL (if not in Registrations) *****/
function _cancelFromWaitlistOnly_(email, classChoice) {
  var waitSheet = _getWaitlistSheet();
  var wVals = waitSheet.getDataRange().getValues();
  if (!wVals || wVals.length < 2) return { success:false, error:'No matching registration found to cancel.' };

  var wmap = _headerMap(wVals[0] || []);
  var cEmail = _idx(wmap, ['email','e-mail']);
  var cClass = _idx(wmap, ['class choice','class','classchoice']);
  var cPos   = _idx(wmap, ['position']);
  if (cEmail < 0 || cClass < 0) {
    return { success:false, error:'Header mismatch in Waitlist. Expect "Email" & "Class Choice".' };
  }

  var rows = [];
  for (var r = 1; r < wVals.length; r++) {
    var e = _s(wVals[r][cEmail]).toLowerCase();
    var c = _s(wVals[r][cClass]);
    var match = (e === email) && (classChoice ? (c === classChoice) : true);
    if (match) rows.push({ row:r+1, cls:c });
  }

  if (rows.length === 0) return { success:false, error:'No matching registration found to cancel.' };

  for (var i = rows.length - 1; i >= 0; i--) {
    waitSheet.deleteRow(rows[i].row);
  }

  if (cPos >= 0) {
    var affected = {};
    for (i = 0; i < rows.length; i++) affected[rows[i].cls] = true;
    for (var cls in affected) _renumberWaitlist(waitSheet, cls);
  }

  try { _sendEmailCancelled(email, classChoice || null, rows.length); } catch (e6) {}

  var msg = rows.length + ' waitlist entr' + (rows.length === 1 ? 'y' : 'ies') + ' cancelled.';
  return { success:true, deleted:0, removedFromWaitlist:rows.length, promoted:0, message: msg };
}

/***** OPTIONAL: Dashboard safe wrapper *****/
function _safeRebuildDashboard() {
  try { if (typeof _rebuildDashboard === 'function') _rebuildDashboard(); } catch (e) {}
}

/***** SHEET HELPERS *****/
function _getRegistrationSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  for (var i = 0; i < REG_TAB_CANDIDATES.length; i++) {
    var nm = REG_TAB_CANDIDATES[i];
    var sh = ss.getSheetByName(nm);
    if (sh) return _ensureHeaders(sh, ['Timestamp','Name','Email','Phone','Class Choice']);
  }
  var created = ss.insertSheet(REG_TAB_CANDIDATES[0]);
  created.appendRow(['Timestamp','Name','Email','Phone','Class Choice']);
  return created;
}

function _getWaitlistSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(WAIT_TAB_NAME);
  if (!sh) {
    sh = ss.insertSheet(WAIT_TAB_NAME);
    sh.appendRow(['Timestamp','Name','Email','Phone','Class Choice','Position']);
  } else {
    _ensureHeaders(sh, ['Timestamp','Name','Email','Phone','Class Choice','Position']);
  }
  return sh;
}

function _ensureHeaders(sh, headers) {
  if (sh.getLastRow() === 0) { sh.appendRow(headers); return sh; }
  var first = sh.getRange(1, 1, 1, Math.max(headers.length, 1)).getValues()[0] || [];
  var ok = true;
  for (var i = 0; i < headers.length; i++) if (first[i] !== headers[i]) { ok = false; break; }
  if (!ok) { sh.insertRows(1); sh.getRange(1, 1, 1, headers.length).setValues([headers]); }
  return sh;
}

/***** EMAILS — bilingual (EN + FR) and 12-hour policy *****/
function _sendEmailRegistered(email, name, classChoiceDisplay) {
  if (!email) return;
  var subject = 'uOttawa Boxing Club: Registration Confirmed | Inscription confirmée';
  var body =
    'Hi ' + (name || '') + ',\n\n' +
    'You are registered for: ' + classChoiceDisplay + '\n\n' +
    'Please arrive 10 minutes early to wrap up and warm up.\n' +
    'Cancellation policy: cancellations are NOT allowed within 12 hours of class time.\n' +
    'Late arrivals may lose their spot to waitlisted members.\n\n' +
    'See you in the gym! 🥊\n' +
    'uOttawa Boxing Club Team\n' +
    '\n' +
    '— — — — — — — — — — — —\n' +
    'Bonjour ' + (name || '') + ',\n\n' +
    'Votre inscription est confirmée pour : ' + classChoiceDisplay + '\n\n' +
    'Veuillez arriver 10 minutes à l’avance pour vous préparer et vous échauffer.\n' +
    'Politique d’annulation : les annulations ne sont PAS permises dans les 12 heures précédant le cours.\n' +
    'Les retards peuvent entraîner la perte de votre place au profit d’une personne sur la liste d’attente.\n\n' +
    'À bientôt au gym! 🥊\n' +
    'Équipe du Club de boxe de l’uOttawa';
  MailApp.sendEmail(email, subject, body);
}

function _sendEmailWaitlisted(email, name, classChoiceDisplay, position) {
  if (!email) return;
  var subject = 'uOttawa Boxing Club: You’re on the Waitlist | Vous êtes sur la liste d’attente';
  var body =
    'Hi ' + (name || '') + ',\n\n' +
    classChoiceDisplay + ' is currently full. Your waitlist position is ' + position + '.\n' +
    'We will email you automatically if a spot opens.\n' +
    'Cancellation policy (for confirmed spots): no cancellations within 12 hours of class time.\n\n' +
    'Thanks for your patience,\n' +
    'uOttawa Boxing Club Team\n' +
    '\n' +
    '— — — — — — — — — — — —\n' +
    'Bonjour ' + (name || '') + ',\n\n' +
    classChoiceDisplay + ' est actuellement complet. Votre position sur la liste d’attente est ' + position + '.\n' +
    'Nous vous enverrons un courriel automatiquement si une place se libère.\n' +
    'Politique d’annulation (pour les places confirmées) : aucune annulation dans les 12 heures précédant le cours.\n\n' +
    'Merci de votre patience,\n' +
    'Équipe du Club de boxe de l’uOttawa';
  MailApp.sendEmail(email, subject, body);
}

function _sendEmailPromoted(email, name, classChoiceDisplay) {
  if (!email) return;
  var subject = 'uOttawa Boxing Club: A Spot Opened — You’re In! | Une place s’est libérée — vous êtes inscrit(e)!';
  var body =
    'Hi ' + (name || '') + ',\n\n' +
    'Good news! A spot opened for: ' + classChoiceDisplay + '\n' +
    'You’ve been moved from the waitlist to registered.\n' +
    'Please arrive 10 minutes early. Cancellation policy: no cancellations within 12 hours of class time.\n\n' +
    'uOttawa Boxing Club Team\n' +
    '\n' +
    '— — — — — — — — — — — —\n' +
    'Bonjour ' + (name || '') + ',\n\n' +
    'Bonne nouvelle! Une place s’est libérée pour : ' + classChoiceDisplay + '\n' +
    'Vous avez été déplacé(e) de la liste d’attente vers la liste des personnes inscrites.\n' +
    'Veuillez arriver 10 minutes à l’avance. Politique d’annulation : aucune annulation dans les 12 heures précédant le cours.\n\n' +
    'Équipe du Club de boxe de l’uOttawa';
  MailApp.sendEmail(email, subject, body);
}

function _sendEmailCancelled(email, classChoice, count) {
  if (!email) return;
  var display = classChoice ? _displaySession_(classChoice) : '';
  var subject = 'uOttawa Boxing Club: Cancellation Received | Annulation reçue';
  var body =
    'Hi,\n\n' +
    'We cancelled ' + count + ' registration' + (count === 1 ? '' : 's') +
    (display ? ' for ' + display : '') + '.\n' +
    'Reminder: cancellations are not permitted within 12 hours of class time.\n\n' +
    'If this wasn’t you, please reply to let us know.\n\n' +
    'uOttawa Boxing Club Team\n' +
    '\n' +
    '— — — — — — — — — — — —\n' +
    'Bonjour,\n\n' +
    'Nous avons annulé ' + count + ' inscription' + (count === 1 ? '' : 's') +
    (display ? ' pour ' + display : '') + '.\n' +
    'Rappel : les annulations ne sont pas permises dans les 12 heures précédant le cours.\n\n' +
    'Si ce n’était pas vous, veuillez répondre à ce courriel pour nous en informer.\n\n' +
    'Équipe du Club de boxe de l’uOttawa';
  MailApp.sendEmail(email, subject, body);
}

/***** ROBUST HEADER HELPERS *****/
function _normalizeHeader(h) {
  return (h == null ? '' : String(h)).replace(/\u00A0/g, ' ').trim().toLowerCase();
}
function _headerMap(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) map[_normalizeHeader(headers[i])] = i;
  return map;
}
function _idx(map, names) {
  for (var i = 0; i < names.length; i++) {
    var k = _normalizeHeader(names[i]);
    if (map.hasOwnProperty(k)) return map[k];
  }
  return -1;
}

/***** UTILS *****/
function _s(v) { return (v == null ? '' : String(v)).trim(); }

/***** DIAGNOSTICS *****/
function pingSetup() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var names = ss.getSheets().map(function(s){ return s.getName(); });
  var reg = _getRegistrationSheet();
  var wait = _getWaitlistSheet();
  return {
    spreadsheetName: ss.getName(),
    tabsFound: names,
    usingRegistrationTab: reg.getName(),
    usingWaitlistTab: wait.getName(),
    regRows: reg.getLastRow(),
    waitRows: wait.getLastRow()
  };
}

function testEmailOnce() {
  var me = Session.getActiveUser().getEmail() || 'your-email@example.com';
  MailApp.sendEmail(me, 'uOttawa Boxing Club — Test | Essai', 'If you see this, email sending is authorized.\n\nSi vous voyez ceci, l’envoi de courriels est autorisé.');
  return 'Sent test to: ' + me;
}
