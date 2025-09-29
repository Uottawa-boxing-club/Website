/***** CONFIG *****/
var SHEET_ID = '1HyoKB0AxA3plMrFhzcYze7JmhrHZEawbjvwND-vUaFM';

// Auto-detect these registration tab names, in order:
var REG_TAB_CANDIDATES = ['Registrations','Sheet1','Registration','Signups'];
// Waitlist tab name:
var WAIT_TAB_NAME = 'Waitlist';
// Dashboard tab name (if you kept the dashboard feature; harmless if absent)
var DASHBOARD_TAB_NAME = 'Dashboard';

// Seats per class
var CAPACITY_PER_CLASS = 22;

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
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // allow embedding in iframe
}

/***** TEMPORARY SCHEDULE OVERRIDE HELPERS *****/
// Next week only (Ottawa time): Mon Sep 29 → Mon Oct 6 (exclusive)
// Make Tue/Wed active NOW through Oct 6 (exclusive)
function _isOverrideWeek_(now) {
  var today = now || new Date();
  var start = new Date('2025-09-27T00:00:00-04:00'); // <-- moved earlier so it applies today
  var end   = new Date('2025-10-06T00:00:00-04:00'); // Mon, Oct 6, 2025 (exclusive)
  return today >= start && today < end;
}


// Returns labels & regex for grouping classes on the dashboard
function _currentSchedule_() {
  if (_isOverrideWeek_()) {
    // NEXT WEEK ONLY: Tuesday + Wednesday
    return {
      day1: { name: 'TUESDAY',   headerReg: /tuesday/i,   sectionLabel: 'TUESDAY 4:30-5:30 PM' },
      day2: { name: 'WEDNESDAY', headerReg: /wednesday/i, sectionLabel: 'WEDNESDAY 4:00-5:00 PM' }
    };
  }
  // DEFAULT: Tuesday + Friday
  return {
    day1: { name: 'TUESDAY', headerReg: /tuesday/i, sectionLabel: 'TUESDAY 4:30-5:30 PM' },
    day2: { name: 'FRIDAY',  headerReg: /friday/i,  sectionLabel: 'FRIDAY 2:30-3:30 PM' }
  };
}

/***** MAIN ACTIONS *****/
function registerNew(form) {
  var name = _s(form && form.name);
  var email = _s(form && form.email).toLowerCase();
  var phone = _s(form && form.phone);
  var classChoice = _s(form && form.classChoice);

  if (!name || !email || !phone || !classChoice) {
    return { success: false, error: 'Please provide name, email, phone, and class day.' };
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

  // Prevent duplicate (same email + class)
  for (var i = 1; i < regData.length; i++) {
    var e = _s(regData[i][cEmail]).toLowerCase();
    var c = _s(regData[i][cClass]);
    if (e === email && c === classChoice) {
      try { _sendEmailRegistered(email, name, classChoice); } catch (e1) {}
      _safeRebuildDashboard();
      return { success:true, status:'registered', message:'You are already registered for this class.' };
    }
  }

  // Count active in this class
  var active = 0;
  for (var j = 1; j < regData.length; j++) {
    var c2 = _s(regData[j][cClass]);
    if (c2 === classChoice) active++;
  }

  if (active < CAPACITY_PER_CLASS) {
    regSheet.appendRow([new Date(), name, email, phone, classChoice]);
    try { _sendEmailRegistered(email, name, classChoice); } catch (e2) {}
    _safeRebuildDashboard();
    return { success:true, status:'registered', message:'Registered successfully.' };
  }

  // Waitlist (position within this class)
  var waitData = waitSheet.getDataRange().getValues();
  var wmap = _headerMap(waitData[0] || []);
  var wClass = _idx(wmap, ['class choice','class','classchoice']);
  var position = 1;
  for (var k = 1; k < waitData.length; k++) {
    var wc = _s(waitData[k][wClass]);
    if (wc === classChoice) position++;
  }
  waitSheet.appendRow([new Date(), name, email, phone, classChoice, position]);
  try { _sendEmailWaitlisted(email, name, classChoice, position); } catch (e3) {}
  _safeRebuildDashboard();
  return {
    success:true, status:'waitlisted', position:position,
    message:'Class is full. You are on the waitlist (position ' + position + ').'
  };
}

function cancelRegistration(form) {
  var email = _s(form && form.email).toLowerCase();
  var classChoice = _s(form && form.classChoice); // optional

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

  try { _sendEmailPromoted(email, name, classChoice); } catch (e5) {}
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
function _sendEmailRegistered(email, name, classChoice) {
  if (!email) return;
  var subject = 'uOttawa Boxing Club: Registration Confirmed | Inscription confirmée';
  var body =
    'Hi ' + (name || '') + ',\n\n' +
    'You are registered for: ' + classChoice + '\n\n' +
    'Please arrive 10 minutes early to wrap up and warm up.\n' +
    'Cancellation policy: cancellations are NOT allowed within 12 hours of class time.\n' +
    'Late arrivals may lose their spot to waitlisted members.\n\n' +
    'See you in the gym! 🥊\n' +
    'uOttawa Boxing Club Team\n' +
    '\n' +
    '— — — — — — — — — — — —\n' +
    'Bonjour ' + (name || '') + ',\n\n' +
    'Votre inscription est confirmée pour : ' + classChoice + '\n\n' +
    'Veuillez arriver 10 minutes à l’avance pour vous préparer et vous échauffer.\n' +
    'Politique d’annulation : les annulations ne sont PAS permises dans les 12 heures précédant le cours.\n' +
    'Les retards peuvent entraîner la perte de votre place au profit d’une personne sur la liste d’attente.\n\n' +
    'À bientôt au gym! 🥊\n' +
    'Équipe du Club de boxe de l’uOttawa';
  MailApp.sendEmail(email, subject, body);
}

function _sendEmailWaitlisted(email, name, classChoice, position) {
  if (!email) return;
  var subject = 'uOttawa Boxing Club: You’re on the Waitlist | Vous êtes sur la liste d’attente';
  var body =
    'Hi ' + (name || '') + ',\n\n' +
    classChoice + ' is currently full. Your waitlist position is ' + position + '.\n' +
    'We will email you automatically if a spot opens.\n' +
    'Cancellation policy (for confirmed spots): no cancellations within 12 hours of class time.\n\n' +
    'Thanks for your patience,\n' +
    'uOttawa Boxing Club Team\n' +
    '\n' +
    '— — — — — — — — — — — —\n' +
    'Bonjour ' + (name || '') + ',\n\n' +
    classChoice + ' est actuellement complet. Votre position sur la liste d’attente est ' + position + '.\n' +
    'Nous vous enverrons un courriel automatiquement si une place se libère.\n' +
    'Politique d’annulation (pour les places confirmées) : aucune annulation dans les 12 heures précédant le cours.\n\n' +
    'Merci de votre patience,\n' +
    'Équipe du Club de boxe de l’uOttawa';
  MailApp.sendEmail(email, subject, body);
}

function _sendEmailPromoted(email, name, classChoice) {
  if (!email) return;
  var subject = 'uOttawa Boxing Club: A Spot Opened — You’re In! | Une place s’est libérée — vous êtes inscrit(e)!';
  var body =
    'Hi ' + (name || '') + ',\n\n' +
    'Good news! A spot opened for: ' + classChoice + '\n' +
    'You’ve been moved from the waitlist to registered.\n' +
    'Please arrive 10 minutes early. Cancellation policy: no cancellations within 12 hours of class time.\n\n' +
    'uOttawa Boxing Club Team\n' +
    '\n' +
    '— — — — — — — — — — — —\n' +
    'Bonjour ' + (name || '') + ',\n\n' +
    'Bonne nouvelle! Une place s’est libérée pour : ' + classChoice + '\n' +
    'Vous avez été déplacé(e) de la liste d’attente vers la liste des personnes inscrites.\n' +
    'Veuillez arriver 10 minutes à l’avance. Politique d’annulation : aucune annulation dans les 12 heures précédant le cours.\n\n' +
    'Équipe du Club de boxe de l’uOttawa';
  MailApp.sendEmail(email, subject, body);
}

function _sendEmailCancelled(email, classChoice, count) {
  if (!email) return;
  var subject = 'uOttawa Boxing Club: Cancellation Received | Annulation reçue';
  var body =
    'Hi,\n\n' +
    'We cancelled ' + count + ' registration' + (count === 1 ? '' : 's') +
    (classChoice ? ' for ' + classChoice : '') + '.\n' +
    'Reminder: cancellations are not permitted within 12 hours of class time.\n\n' +
    'If this wasn’t you, please reply to let us know.\n\n' +
    'uOttawa Boxing Club Team\n' +
    '\n' +
    '— — — — — — — — — — — —\n' +
    'Bonjour,\n\n' +
    'Nous avons annulé ' + count + ' inscription' + (count === 1 ? '' : 's') +
    (classChoice ? ' pour ' + classChoice : '') + '.\n' +
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
/***** MANUAL RESET FUNCTION *****/
function clearAllRegistrations() {
  var regSheet = _getRegistrationSheet();
  var waitSheet = _getWaitlistSheet();
  
  // Clear all data except headers
  if (regSheet.getLastRow() > 1) {
    regSheet.deleteRows(2, regSheet.getLastRow() - 1);
  }
  if (waitSheet.getLastRow() > 1) {
    waitSheet.deleteRows(2, waitSheet.getLastRow() - 1);
  }
  
  // Clear dashboard sheet if it exists
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var dashSheet = ss.getSheetByName(DASHBOARD_TAB_NAME);
  if (dashSheet) {
    dashSheet.clear();
  }
  _safeRebuildDashboard();
  return 'Reset complete. All registrations and dashboard cleared.';
}

/*** DASHBOARD (DYNAMIC SCHEDULE) ***/
function setupDashboardHeaders() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var dashSheet = ss.getSheetByName(DASHBOARD_TAB_NAME);
  if (!dashSheet) dashSheet = ss.insertSheet(DASHBOARD_TAB_NAME);
  dashSheet.clear();
  dashSheet.appendRow(['uOttawa Boxing Club - Registration Dashboard']);
  dashSheet.appendRow(['Last Updated: ' + new Date().toString()]);
  dashSheet.appendRow(['']);
  dashSheet.getRange(1, 1).setFontWeight('bold').setFontSize(16);
  dashSheet.autoResizeColumns(1, 5);
  return 'Clean dashboard created successfully';
}

function _rebuildDashboard() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var dashSheet = ss.getSheetByName(DASHBOARD_TAB_NAME);
    if (!dashSheet || dashSheet.getLastRow() < 3) {
      setupDashboardHeaders();
      dashSheet = ss.getSheetByName(DASHBOARD_TAB_NAME);
    }

    dashSheet.getRange(2, 1).setValue('Last Updated: ' + new Date().toString());

    var lastRow = dashSheet.getLastRow();
    if (lastRow > 3) dashSheet.deleteRows(4, lastRow - 3);

    // Current schedule (override next week only)
    var sched = _currentSchedule_();

    var regSheet = _getRegistrationSheet();
    var waitSheet = _getWaitlistSheet();
    var regData = regSheet.getDataRange().getValues();
    var waitData = waitSheet.getDataRange().getValues();

    var d1Regs = [], d2Regs = [];
    var d1Wait = [], d2Wait = [];

    if (regData.length > 1) {
      var rmap  = _headerMap(regData[0] || []);
      var cName = _idx(rmap, ['name']);
      var cEmail= _idx(rmap, ['email','e-mail']);
      var cPhone= _idx(rmap, ['phone']);
      var cClass= _idx(rmap, ['class choice','class','classchoice']);
      var cTime = _idx(rmap, ['timestamp']);

      for (var i = 1; i < regData.length; i++) {
        var raw = _s(regData[i][cClass]);
        var info = [regData[i][cName]||'', regData[i][cEmail]||'', regData[i][cPhone]||'', regData[i][cTime]||''];
        if (sched.day1.headerReg.test(raw)) d1Regs.push(info);
        else if (sched.day2.headerReg.test(raw)) d2Regs.push(info);
      }
    }

    if (waitData.length > 1) {
      var wmap  = _headerMap(waitData[0] || []);
      var wName = _idx(wmap, ['name']);
      var wEmail= _idx(wmap, ['email','e-mail']);
      var wPhone= _idx(wmap, ['phone']);
      var wClass= _idx(wmap, ['class choice','class','classchoice']);
      var wPos  = _idx(wmap, ['position']);
      var wTime = _idx(wmap, ['timestamp']);

      for (var j = 1; j < waitData.length; j++) {
        var raw = _s(waitData[j][wClass]);
        var winfo = [waitData[j][wPos]||'', waitData[j][wName]||'', waitData[j][wEmail]||'', waitData[j][wPhone]||'', waitData[j][wTime]||''];
        if (sched.day1.headerReg.test(raw)) d1Wait.push(winfo);
        else if (sched.day2.headerReg.test(raw)) d2Wait.push(winfo);
      }
    }

    var all = [];

    // DAY 1
    all.push([sched.day1.sectionLabel + ' REGISTRATIONS (' + d1Regs.length + '/' + CAPACITY_PER_CLASS + ')']);
    all.push(['Name','Email','Phone','Registration Time']);
    if (d1Regs.length) d1Regs.forEach(function(r){ all.push(r); }); else all.push(['No registrations yet']);

    all.push(['']);
    all.push([sched.day1.sectionLabel + ' WAITLIST (' + d1Wait.length + ')']);
    all.push(['Position','Name','Email','Phone','Waitlist Time']);
    if (d1Wait.length) d1Wait.forEach(function(r){ all.push(r); }); else all.push(['No waitlist']);

    // DAY 2
    all.push(['']);
    all.push([sched.day2.sectionLabel + ' REGISTRATIONS (' + d2Regs.length + '/' + CAPACITY_PER_CLASS + ')']);
    all.push(['Name','Email','Phone','Registration Time']);
    if (d2Regs.length) d2Regs.forEach(function(r){ all.push(r); }); else all.push(['No registrations yet']);

    all.push(['']);
    all.push([sched.day2.sectionLabel + ' WAITLIST (' + d2Wait.length + ')']);
    all.push(['Position','Name','Email','Phone','Waitlist Time']);
    if (d2Wait.length) d2Wait.forEach(function(r){ all.push(r); }); else all.push(['No waitlist']);

    // pad to 5 columns
    for (var a = 0; a < all.length; a++) while (all[a].length < 5) all[a].push('');

    dashSheet.getRange(4, 1, all.length, 5).setValues(all);

    // Format headers
    var startRow = 4;
    for (var idx = 0; idx < all.length; idx++) {
      var v = (all[idx][0] || '').toString();
      var row = startRow + idx;
      if (v.indexOf('REGISTRATIONS (') > -1 && v.indexOf(sched.day1.name) === 0) {
        dashSheet.getRange(row, 1, 1, 5).setFontWeight('bold').setFontSize(12).setBackground('#e3f2fd');
      } else if (v.indexOf('WAITLIST (') > -1 && v.indexOf(sched.day1.name) === 0) {
        dashSheet.getRange(row, 1, 1, 5).setFontWeight('bold').setFontSize(12).setBackground('#fff3e0');
      } else if (v.indexOf('REGISTRATIONS (') > -1 && v.indexOf(sched.day2.name) === 0) {
        dashSheet.getRange(row, 1, 1, 5).setFontWeight('bold').setFontSize(12).setBackground('#e8f5e8');
      } else if (v.indexOf('WAITLIST (') > -1 && v.indexOf(sched.day2.name) === 0) {
        dashSheet.getRange(row, 1, 1, 5).setFontWeight('bold').setFontSize(12).setBackground('#fce4ec');
      } else if (v === 'Name' || v === 'Position') {
        dashSheet.getRange(row, 1, 1, 5).setFontWeight('bold').setBackground('#f5f5f5');
      }
    }

    dashSheet.autoResizeColumns(1, 5);

  } catch (error) {
    Logger.log('Dashboard rebuild failed: ' + error.toString());
    try { setupDashboardHeaders(); } catch (e2) {}
  }
}

/*** DASHBOARD AUTO-UPDATE TRIGGER SETUP ***/
function setupDashboardTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === '_rebuildDashboard') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('_rebuildDashboard').timeBased().everyMinutes(10).create();
  return 'Dashboard auto-update trigger installed (every 10 minutes)';
}

/*** WEEKLY RESET TRIGGER SETUP ***/
function setupWeeklyResetTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'clearAllRegistrations') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('clearAllRegistrations')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY)
    .atHour(11)
    .nearMinute(30)
    .inTimezone('America/Toronto')
    .create();
  return 'Weekly reset trigger installed for Saturday at 11:30 AM Ottawa time';
}
