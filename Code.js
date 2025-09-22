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
    // English
    'Hi ' + (name || '') + ',\n\n' +
    'You are registered for: ' + classChoice + '\n\n' +
    'Please arrive 10 minutes early to wrap up and warm up.\n' +
    'Cancellation policy: cancellations are NOT allowed within 12 hours of class time.\n' +
    'Late arrivals may lose their spot to waitlisted members.\n\n' +
    'See you in the gym! 🥊\n' +
    'uOttawa Boxing Club Team\n' +
    '\n' +
    '— — — — — — — — — — — —\n' +
    // French
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
    // English
    'Hi ' + (name || '') + ',\n\n' +
    classChoice + ' is currently full. Your waitlist position is ' + position + '.\n' +
    'We will email you automatically if a spot opens.\n' +
    'Cancellation policy (for confirmed spots): no cancellations within 12 hours of class time.\n\n' +
    'Thanks for your patience,\n' +
    'uOttawa Boxing Club Team\n' +
    '\n' +
    '— — — — — — — — — — — —\n' +
    // French
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
    // English
    'Hi ' + (name || '') + ',\n\n' +
    'Good news! A spot opened for: ' + classChoice + '\n' +
    'You’ve been moved from the waitlist to registered.\n' +
    'Please arrive 10 minutes early. Cancellation policy: no cancellations within 12 hours of class time.\n\n' +
    'uOttawa Boxing Club Team\n' +
    '\n' +
    '— — — — — — — — — — — —\n' +
    // French
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
    // English
    'Hi,\n\n' +
    'We cancelled ' + count + ' registration' + (count === 1 ? '' : 's') +
    (classChoice ? ' for ' + classChoice : '') + '.\n' +
    'Reminder: cancellations are not permitted within 12 hours of class time.\n\n' +
    'If this wasn’t you, please reply to let us know.\n\n' +
    'uOttawa Boxing Club Team\n' +
    '\n' +
    '— — — — — — — — — — — —\n' +
    // French
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
    Logger.log('Dashboard cleared');
  }
  
  // Rebuild dashboard if it exists
  _safeRebuildDashboard();
  
  Logger.log('All registrations, waitlist entries, and dashboard cleared at: ' + new Date());
  return 'Reset complete. All registrations and dashboard cleared.';
}
/***** CLEAN 4-SECTION DASHBOARD FUNCTIONS - FIXED SPACING *****/

function setupDashboardHeaders() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var dashSheet = ss.getSheetByName(DASHBOARD_TAB_NAME);
  
  // Create dashboard sheet if it doesn't exist
  if (!dashSheet) {
    dashSheet = ss.insertSheet(DASHBOARD_TAB_NAME);
  }
  
  // Clear everything
  dashSheet.clear();
  
  // Header only
  dashSheet.appendRow(['uOttawa Boxing Club - Registration Dashboard']);
  dashSheet.appendRow(['Last Updated: ' + new Date().toString()]);
  dashSheet.appendRow(['']);
  
  // Format headers
  dashSheet.getRange(1, 1).setFontWeight('bold').setFontSize(16);
  
  // Auto-resize columns
  dashSheet.autoResizeColumns(1, 5);
  
  Logger.log('Clean dashboard headers set up at: ' + new Date());
  return 'Clean dashboard created successfully';
}

function _rebuildDashboard() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var dashSheet = ss.getSheetByName(DASHBOARD_TAB_NAME);
    
    // If dashboard doesn't exist or is too small, recreate it
    if (!dashSheet || dashSheet.getLastRow() < 3) {
      setupDashboardHeaders();
      dashSheet = ss.getSheetByName(DASHBOARD_TAB_NAME);
    }
    
    // Update timestamp
    dashSheet.getRange(2, 1).setValue('Last Updated: ' + new Date().toString());
    
    // Clear everything below the header (row 4 onwards)
    var lastRow = dashSheet.getLastRow();
    if (lastRow > 3) {
      dashSheet.deleteRows(4, lastRow - 3);
    }
    
    // Get data using existing helper functions
    var regSheet = _getRegistrationSheet();
    var waitSheet = _getWaitlistSheet();
    var regData = regSheet.getDataRange().getValues();
    var waitData = waitSheet.getDataRange().getValues();
    
    // Process registration data
    var tuesdayRegs = [];
    var fridayRegs = [];
    
    if (regData.length > 1) {
      var rmap = _headerMap(regData[0] || []);
      var cName = _idx(rmap, ['name']);
      var cEmail = _idx(rmap, ['email','e-mail']);
      var cPhone = _idx(rmap, ['phone']);
      var cClass = _idx(rmap, ['class choice','class','classchoice']);
      var cTime = _idx(rmap, ['timestamp']);
      
      for (var i = 1; i < regData.length; i++) {
        var classChoiceRaw = _s(regData[i][cClass]);
        var classChoice = classChoiceRaw.replace(/\s+/g, ' ').trim().toLowerCase();
        var regInfo = [
          regData[i][cName] || '',
          regData[i][cEmail] || '',
          regData[i][cPhone] || '',
          regData[i][cTime] || ''
        ];
        // Bulletproof matching for class names
        if (/tuesday/i.test(classChoiceRaw)) {
          tuesdayRegs.push(regInfo);
        } else if (/friday/i.test(classChoiceRaw)) {
          fridayRegs.push(regInfo);
        }
      }
    }
    
    // Process waitlist data
    var tuesdayWait = [];
    var fridayWait = [];
    
    if (waitData.length > 1) {
      var wmap = _headerMap(waitData[0] || []);
      var wName = _idx(wmap, ['name']);
      var wEmail = _idx(wmap, ['email','e-mail']);
      var wPhone = _idx(wmap, ['phone']);
      var wClass = _idx(wmap, ['class choice','class','classchoice']);
      var wPos = _idx(wmap, ['position']);
      var wTime = _idx(wmap, ['timestamp']);
      
      for (var j = 1; j < waitData.length; j++) {
        var classChoiceRaw = _s(waitData[j][wClass]);
        var classChoice = classChoiceRaw.replace(/\s+/g, ' ').trim().toLowerCase();
        var waitInfo = [
          waitData[j][wPos] || '',
          waitData[j][wName] || '',
          waitData[j][wEmail] || '',
          waitData[j][wPhone] || '',
          waitData[j][wTime] || ''
        ];
        // Bulletproof matching for class names
        if (/tuesday/i.test(classChoiceRaw)) {
          tuesdayWait.push(waitInfo);
        } else if (/friday/i.test(classChoiceRaw)) {
          fridayWait.push(waitInfo);
        }
      }
    }
    
    // Build all content in arrays first, then write to sheet at once
    var allContent = [];
    
    // SECTION 1: TUESDAY REGISTRATIONS
    allContent.push(['TUESDAY 4:30-5:30 PM REGISTRATIONS (' + tuesdayRegs.length + '/' + CAPACITY_PER_CLASS + ')']);
    allContent.push(['Name', 'Email', 'Phone', 'Registration Time']);
    
    if (tuesdayRegs.length > 0) {
      for (var t = 0; t < tuesdayRegs.length; t++) {
        allContent.push(tuesdayRegs[t]);
      }
    } else {
      allContent.push(['No registrations yet']);
    }
    
    // SECTION 2: TUESDAY WAITLIST
    allContent.push(['']); // One blank row for spacing
    allContent.push(['TUESDAY 4:30-5:30 PM WAITLIST (' + tuesdayWait.length + ')']);  
    allContent.push(['Position', 'Name', 'Email', 'Phone', 'Waitlist Time']);
    
    
    if (tuesdayWait.length > 0) {
      for (var tw = 0; tw < tuesdayWait.length; tw++) {
        allContent.push(tuesdayWait[tw]);
      }
    } else {
      allContent.push(['No waitlist']);
    }
    
    // SECTION 3: FRIDAY REGISTRATIONS
    allContent.push(['']); // One blank row for spacing
    allContent.push(['FRIDAY 2:30-3:30 PM REGISTRATIONS (' + fridayRegs.length + '/' + CAPACITY_PER_CLASS + ')']);
    allContent.push(['Name', 'Email', 'Phone', 'Registration Time']);
    
    if (fridayRegs.length > 0) {
      for (var f = 0; f < fridayRegs.length; f++) {
        allContent.push(fridayRegs[f]);
      }
    } else {
      allContent.push(['No registrations yet']);
    }
    
    // SECTION 4: FRIDAY WAITLIST
    allContent.push(['']); // One blank row for spacing
    allContent.push(['FRIDAY 2:30-3:30 PM WAITLIST (' + fridayWait.length + ')']);
    allContent.push(['Position', 'Name', 'Email', 'Phone', 'Waitlist Time']);
    
    if (fridayWait.length > 0) {
      for (var fw = 0; fw < fridayWait.length; fw++) {
        allContent.push(fridayWait[fw]);
      }
    } else {
      allContent.push(['No waitlist']);
    }
    
    // Write all content at once starting from row 4
    if (allContent.length > 0) {
      // Ensure all rows have the same number of columns (5)
      for (var c = 0; c < allContent.length; c++) {
        while (allContent[c].length < 5) {
          allContent[c].push('');
        }
      }
      dashSheet.getRange(4, 1, allContent.length, 5).setValues(allContent);
    }
    
    // Format section headers - track row positions
    var currentRow = 4;
    var sectionHeaders = [];
    
    for (var a = 0; a < allContent.length; a++) {
      var cellValue = allContent[a][0] ? allContent[a][0].toString() : '';
      
      if (cellValue.includes('TUESDAY') && cellValue.includes('REGISTRATIONS')) {
        sectionHeaders.push({row: currentRow + a, type: 'tues_reg'});
      } else if (cellValue.includes('TUESDAY') && cellValue.includes('WAITLIST')) {
        sectionHeaders.push({row: currentRow + a, type: 'tues_wait'});
      } else if (cellValue.includes('FRIDAY') && cellValue.includes('REGISTRATIONS')) {
        sectionHeaders.push({row: currentRow + a, type: 'fri_reg'});
      } else if (cellValue.includes('FRIDAY') && cellValue.includes('WAITLIST')) {
        sectionHeaders.push({row: currentRow + a, type: 'fri_wait'});
      } else if (cellValue === 'Name' || cellValue === 'Position') {
        sectionHeaders.push({row: currentRow + a, type: 'column_header'});
      }
    }
    
    // Apply formatting with different colors for each section
    for (var h = 0; h < sectionHeaders.length; h++) {
      var header = sectionHeaders[h];
      switch(header.type) {
        case 'tues_reg':
          dashSheet.getRange(header.row, 1, 1, 5).setFontWeight('bold').setFontSize(12).setBackground('#e3f2fd'); // Light Blue
          break;
        case 'tues_wait':
          dashSheet.getRange(header.row, 1, 1, 5).setFontWeight('bold').setFontSize(12).setBackground('#fff3e0'); // Light Orange
          break;
        case 'fri_reg':
          dashSheet.getRange(header.row, 1, 1, 5).setFontWeight('bold').setFontSize(12).setBackground('#e8f5e8'); // Light Green
          break;
        case 'fri_wait':
          dashSheet.getRange(header.row, 1, 1, 5).setFontWeight('bold').setFontSize(12).setBackground('#fce4ec'); // Light Pink
          break;
        case 'column_header':
          dashSheet.getRange(header.row, 1, 1, 5).setFontWeight('bold').setBackground('#f5f5f5'); // Light Gray
          break;
      }
    }
    
    // Auto-resize columns
    dashSheet.autoResizeColumns(1, 5);
    
    Logger.log('Clean 4-section dashboard rebuilt successfully at: ' + new Date());
    
  } catch (error) {
    Logger.log('Dashboard rebuild failed: ' + error.toString());
    try {
      setupDashboardHeaders();
      Logger.log('Dashboard recreated due to error');
    } catch (e2) {
      Logger.log('Failed to recreate dashboard: ' + e2.toString());
    }
  }
}

// Test function to set up everything
function setupCompleteDashboard() {
  setupDashboardHeaders();
  _rebuildDashboard();
  return 'Complete dashboard setup finished';
}

/*** DASHBOARD AUTO-UPDATE TRIGGER SETUP ***/
function setupDashboardTrigger() {
  // Remove existing triggers for _rebuildDashboard
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === '_rebuildDashboard') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Add a new time-based trigger (every 10 minutes)
  ScriptApp.newTrigger('_rebuildDashboard')
    .timeBased()
    .everyMinutes(10)
    .create();
  return 'Dashboard auto-update trigger installed (every 10 minutes)';
}

/*** WEEKLY RESET TRIGGER SETUP ***/
function setupWeeklyResetTrigger() {
  // Remove existing triggers for clearAllRegistrations
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'clearAllRegistrations') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Ottawa time zone is America/Toronto
  ScriptApp.newTrigger('clearAllRegistrations')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY)
    .atHour(11) // 11 AM
    .nearMinute(30) // 11:30 AM
    .inTimezone('America/Toronto')
    .create();
  return 'Weekly reset trigger installed for Saturday at 11:30 AM Ottawa time';
}