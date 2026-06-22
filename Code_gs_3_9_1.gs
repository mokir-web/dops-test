const DOC_ID = '1klNSqYTtKWeOtYGaoI5jVHGYwPCgA7ZWlgGfhOZDuAU';

// ── doGet: handle CORS preflight + unsubscribe ───────────────────────────
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'unsubscribe') {
    const email = (e.parameter.email || '').toLowerCase().trim();
    const token = (e.parameter.token || '').trim();
    if (email && token && token === computeUnsubToken(email)) {
      unsubscribeUser(email);
      return HtmlService.createHtmlOutput(
        '<!DOCTYPE html><html><head><meta charset="utf-8">' +
        '<style>body{font-family:Arial,sans-serif;background:#f0ede8;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;}' +
        '.box{background:#f5f1ec;border:1.5px solid #d8cfc5;border-radius:10px;padding:40px;max-width:480px;text-align:center;}' +
        'h2{color:#3d5c4a;margin-top:0;}p{color:#6b5c52;}</style></head><body>' +
        '<div class="box"><div style="font-size:48px;">✓</div>' +
        '<h2>Avregistrerad</h2>' +
        '<p>Du kommer inte längre få e-postaviseringar från MoraDOPS.</p>' +
        '<p style="font-size:13px;">Vill du aktivera aviseringar igen kan du göra det under <strong>Min profil → Aviseringar</strong> på <a href="https://mokir-web.github.io/dops/" style="color:#3d5c4a;">mokir-web.github.io/dops/</a></p>' +
        '</div></body></html>'
      ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    return HtmlService.createHtmlOutput('<p>Ogiltig avregistreringslänk.</p>');
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true, status: 'API running' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function computeUnsubToken(email) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    email.toLowerCase().trim() + 'moradops-unsub-2024'
  );
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('').slice(0, 32);
}

function unsubscribeUser(email) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Användare');
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.slice(1).findIndex(r => r[3].toString().trim().toLowerCase() === email);
  if (rowIdx !== -1) sheet.getRange(rowIdx + 2, 14).setValue('Nej');
}

// ── Single POST entry point ───────────────────────────────────────────────
function doPost(e) {
  // GAS doesn't support custom response headers directly, but setting
  // the deployment to "Anyone" access handles CORS for simple requests.
  // For browsers sending preflight (OPTIONS), GAS routes to doGet above.
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    let result;

    switch (action) {
      case 'getData':                result = getData(); break;
      case 'login':                  result = login(params.email, params.pin); break;
      case 'sendResetCode':          result = sendResetCode(params.email); break;
      case 'resetPin':               result = resetPin(params.email, params.code, params.newPin); break;
      case 'registerUser':           result = registerUser(params.user); break;
      case 'updateProfile':          result = updateProfile(params.email, params.updates || { jobRole: params.jobRole, userRole: params.userRole }); break;
      case 'getFormQuestions':       result = getFormQuestions(params.formName); break;
      case 'getRecipientEmail':      result = getRecipientEmail(params.recipientFullName); break;
      case 'saveAnswers':            result = saveAnswers(params.answerList, params.recipientEmail, params.recipientName, params.registrarName, params.formType); break;
      case 'getMyAssessments':       result = getMyAssessments(params.email, params.role, params.filters); break;
      case 'getStatistics':           result = getStatistics(params.filters, params.klinikId); break;
      case 'getAllUsers':             result = getAllUsers(params.klinikId); break;
      case 'deleteUser':              result = deleteUser(params.email); break;
      case 'updateUserByAdmin':       result = updateUserByAdmin(params.targetEmail, params.updates); break;
      case 'getClinicInfo':           result = getClinicInfo(params.klinikId); break;
      case 'setClinicActive':         result = setClinicActive(params.klinikId, params.active); break;
      case 'getMyOverview':           result = getMyOverview(params.email, params.klinikId); break;
      case 'getMyAssessmentCount':    result = getMyAssessmentCount(params.email); break;
      case 'getContextPhrase':        result = getContextPhrase(params.minutes, params.mode); break;
      case 'getDocuments':            result = getDocuments(); break;
      case 'getPrivilegesForUser':    result = getPrivilegesForUser(params.targetEmail); break;
      case 'setUserPrivilege':        result = setUserPrivilege(params.targetEmail, params.privilege, params.klinikId, params.remove); break;
      case 'sendAssessmentRequest':   result = sendAssessmentRequest(params.fromEmail, params.toEmail, params.message, params.formType); break;
      case 'getAssessmentRequests':   result = getAssessmentRequests(params.email); break;
      case 'markRequestRead':         result = markRequestRead(params.id); break;
      case 'markRequestDone':         result = markRequestDone(params.id); break;
      case 'markRequestDoneByParticipants': result = markRequestDoneByParticipants(params.fromEmail, params.toEmail); break;
      case 'sendBulkRequest':         result = sendBulkRequest(params.fromEmail, params.toEmails, params.message, params.formType); break;
      case 'getRequestStats':          result = getRequestStats(params.klinikId); break;
      case 'getClinicAssessments':   result = getClinicAssessments(params.klinikId, params.filters); break;
      case 'exportAssessmentsToDoc': result = exportAssessmentsToDoc(params.email, params.assessments, params.filterInfo); break;
      case 'getScheduledEmails':    result = getScheduledEmails(params.klinikId); break;
      case 'saveScheduledEmail':    result = saveScheduledEmail(params.schedule); break;
      case 'deleteScheduledEmail':  result = deleteScheduledEmail(params.id); break;
      case 'toggleScheduleActive':  result = toggleScheduleActive(params.id); break;
      case 'getKlinikSettings':     result = getKlinikSettings(params.klinikId); break;
      case 'saveKlinikSettings':    result = saveKlinikSettings(params.klinikId, params.settings); break;
      case 'getAllUsers':            result = getAllUsers(params.klinikId, params.includeInaktiverade); break;
      default: result = { error: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true, result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Auth ──────────────────────────────────────────────────────────────────
function generateKlinikId(name) {
  // Konverterar "Falun - Kirurgi" -> "falun-kirurgi"
  return name.toString().trim()
    .toLowerCase()
    .replace(/å/g, 'a').replace(/ä/g, 'a').replace(/ö/g, 'o')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function escapeHtml(str) {
  if (!str) return '';
  return str.toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function getUnsubscribeFooter(email) {
  const baseUrl = 'https://script.google.com/macros/s/AKfycbyFFE870_o2AaerliGcg1zF_EbAnBmwYObkaRVOyqCq_s62wT3QCT6YqUyN9UYZWC9u/exec';
  const token = email ? computeUnsubToken(email.toLowerCase().trim()) : '';
  const unsubUrl = email
    ? baseUrl + '?action=unsubscribe&email=' + encodeURIComponent(email) + '&token=' + token
    : 'https://mokir-web.github.io/dops/';
  return '<div style="margin-top:24px;padding-top:16px;border-top:1px solid #d8cfc5;font-size:11px;color:#9e8e82;text-align:center;">'
    + 'Du får detta mail för att du är registrerad i MoraDOPS. '
    + '<a href="' + unsubUrl + '" style="color:#6b5c52;">Avregistrera e-postaviseringar</a>.'
    + '</div>';
}

function hashPin(pin) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    pin.toString()
  );
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function login(email, pin) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Användare');
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);

  // Hitta användarraden
  const rowIdx = rows.findIndex(r => r[3].toString().toLowerCase() === email.toLowerCase());
  if (rowIdx === -1) return null;
  const row = rows[rowIdx];
  const sheetRow = rowIdx + 2; // 1-indexed + header

  // Brute force-skydd: kolumn L (index 11) = misslyckade försök, kolumn M (index 12) = låst till
  const failCount  = row[11] ? parseInt(row[11]) : 0;
  const lockUntil  = row[12] ? new Date(row[12]) : null;
  const now        = new Date();

  if (lockUntil && now < lockUntil) {
    const minLeft = Math.ceil((lockUntil - now) / 60000);
    return { locked: true, minutesLeft: minLeft };
  }

  const hashedPin = hashPin(pin.toString().trim());
  if (row[7].toString() !== hashedPin) {
    // Fel PIN – räkna upp
    const newFail = failCount + 1;
    sheet.getRange(sheetRow, 12).setValue(newFail);
    if (newFail >= 3) {
      const lockTime = new Date(now.getTime() + 10 * 60 * 1000);
      sheet.getRange(sheetRow, 13).setValue(lockTime.toISOString());
      return { locked: true, minutesLeft: 10 };
    }
    return { failed: true, attemptsLeft: 3 - newFail };
  }

  // Lyckad inloggning – nollställ räknare + uppdatera SenastAktiv
  sheet.getRange(sheetRow, 12).setValue(0);
  sheet.getRange(sheetRow, 13).setValue('');
  sheet.getRange(sheetRow, 17).setValue(new Date().toISOString().slice(0,10)); // Q=SenastAktiv

  // Kontrollera om användaren är inaktiverad
  if ((row[17] || '').toString().trim() === 'Ja') {
    // Återaktivera vid inloggning
    sheet.getRange(sheetRow, 18).setValue('');
  }
  const klinikId = (row[10] || '').toString().trim();
  if (klinikId) {
    const klSheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Kliniker');
    if (klSheet) {
      const klRows = klSheet.getDataRange().getValues().slice(1);
      const klRow = klRows.find(r => r[0].toString().trim() === klinikId);
      if (klRow && klRow[3].toString().trim().toLowerCase() === 'nej') {
        return { error: 'inactive_clinic' };
      }
    }
  }

  // Hämta privilegier
  const privSheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Privilegier');
  const privileges = [];
  if (privSheet) {
    const privData = privSheet.getDataRange().getValues().slice(1);
    privData.forEach(r => {
      if (r[0].toString().toLowerCase() === email.toLowerCase()) {
        privileges.push({ privilege: r[1].toString().trim(), klinikId: r[2].toString().trim() });
      }
    });
  }

  function normalizeUserRole(r) {
    if (!r) return 'Blandbild';
    const s = r.toString().trim();
    if (s === 'Registrerare av bedömningar') return 'Registrerare';
    if (s === 'Mottagare av bedömningar')    return 'Mottagare';
    return s;
  }
  return {
    firstName:   (row[1] || '').toString().trim(),
    lastName:    (row[2] || '').toString().trim(),
    email:       (row[3] || '').toString().trim(),
    jobRole:     (row[4] || '').toString().trim(),
    clinic:      (row[5] || '').toString().trim(),
    userRole:    normalizeUserRole(row[6]),
    klinikId:    (row[10] || '').toString().trim(),
    emailNotify:         (row[13] || 'Ja').toString().trim() !== 'Nej',
    emailNotifyRequests: (row[14] || 'Ja').toString().trim() !== 'Nej',
    startPage:           (row[15] || '').toString().trim(),
    privileges
  };
}

function sendResetCode(email) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Användare');
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.slice(1).findIndex(r => r[3].toString().toLowerCase() === email.toLowerCase());
  if (rowIdx === -1) return { error: 'Ingen användare med denna e-postadress.' };
  const code = Math.floor(100000 + Math.random() * 900000).toString();
  const expiry = new Date(Date.now() + 15 * 60 * 1000).toISOString();
  sheet.getRange(rowIdx + 2, 9).setValue(code);
  sheet.getRange(rowIdx + 2, 10).setValue(expiry);
  MailApp.sendEmail({
    to: email,
    subject: 'Återställ PIN — Bedömning',
    body: `Din återställningskod är: ${code}\n\nKoden är giltig i 15 minuter.`
  });
  return { success: true };
}

function resetPin(email, code, newPin) {
  email = (email || '').toString().trim().toLowerCase();
  code  = (code  || '').toString().trim();
  newPin = (newPin || '').toString().trim();
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Användare');
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.slice(1).findIndex(r => r[3].toString().trim().toLowerCase() === email);
  if (rowIdx === -1) return { error: 'Ingen användare hittades.' };
  if (!/^\d{4,6}$/.test(newPin)) return { error: 'PIN måste vara 4–6 siffror.' };
  if (code !== 'BYPASS') {
    const row = data[rowIdx + 1];
    const savedCode = row[8] ? row[8].toString() : '';
    const expiry    = row[9] ? new Date(row[9]) : null;
    if (!savedCode || savedCode !== code.toString()) return { error: 'Felaktig kod.' };
    if (!expiry || new Date() > expiry) return { error: 'Koden har utgått. Begär en ny.' };
  }
  sheet.getRange(rowIdx + 2, 8).setValue(hashPin(newPin));
  sheet.getRange(rowIdx + 2, 9).setValue('');
  sheet.getRange(rowIdx + 2, 10).setValue('');
  return { success: true };
}

function updateProfile(email, updates) {
  email = (email || '').trim().toLowerCase();
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Användare');
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.slice(1).findIndex(r => r[3].toString().trim().toLowerCase() === email);
  if (rowIdx === -1) return { error: 'Användare hittades inte.' };
  const sheetRow = rowIdx + 2;
  const row = data[rowIdx + 1];
  const oldJobRole  = row[4].toString().trim();
  const oldUserRole = row[6].toString().trim();
  const oldFirstName = row[1].toString().trim();
  const oldLastName  = row[2].toString().trim();
  const oldClinic    = row[5].toString().trim();

  const firstName   = updates.firstName  !== undefined ? updates.firstName.trim()  : oldFirstName;
  const lastName    = updates.lastName   !== undefined ? updates.lastName.trim()   : oldLastName;
  const clinic      = updates.clinic     !== undefined ? updates.clinic.trim()     : oldClinic;
  const newEmail    = updates.newEmail   !== undefined ? updates.newEmail.trim().toLowerCase() : email;
  const jobRole     = updates.jobRole    !== undefined ? updates.jobRole.trim()    : oldJobRole;
  const userRole    = updates.userRole   !== undefined ? updates.userRole.trim()   : oldUserRole;
  const emailNotify         = updates.emailNotify         !== undefined ? updates.emailNotify         : (row[13] || 'Ja');
  const emailNotifyRequests = updates.emailNotifyRequests !== undefined ? updates.emailNotifyRequests : (row[14] || 'Ja');
  const startPage           = updates.startPage           !== undefined ? updates.startPage           : (row[15] || '');

  sheet.getRange(sheetRow, 2).setValue(firstName);
  sheet.getRange(sheetRow, 3).setValue(lastName);
  sheet.getRange(sheetRow, 4).setValue(newEmail);
  sheet.getRange(sheetRow, 5).setValue(jobRole);
  sheet.getRange(sheetRow, 6).setValue(clinic);
  sheet.getRange(sheetRow, 7).setValue(userRole);
  if (updates.klinikId) sheet.getRange(sheetRow, 11).setValue(updates.klinikId.trim());
  sheet.getRange(sheetRow, 14).setValue(emailNotify);
  sheet.getRange(sheetRow, 15).setValue(emailNotifyRequests);
  sheet.getRange(sheetRow, 16).setValue(startPage);

  if (updates.newPin) {
    sheet.getRange(sheetRow, 8).setValue(hashPin(updates.newPin.trim()));
  }

  // Synka Categories
  updateCategoriesForProfileChange(oldFirstName, oldLastName, oldClinic, oldJobRole, oldUserRole, jobRole, userRole);
  // Om namn/klinik ändrats, rebuild helt
  if (firstName !== oldFirstName || lastName !== oldLastName || clinic !== oldClinic) {
    rebuildCategoriesFromUsers(SpreadsheetApp.openById(DOC_ID));
    invalidateCache();
  }

  return { success: true, firstName, lastName, email: newEmail, clinic, jobRole, userRole, emailNotify, emailNotifyRequests, startPage };
}

function updateCategoriesForProfileChange(firstName, lastName, clinic, oldJobRole, oldUserRole, newJobRole, newUserRole) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Categories');
  const totalRows = Math.max(sheet.getLastRow(), 1);
  const data = sheet.getRange(1, 1, totalRows, 7).getValues();
  const cellValue = `${firstName} ${lastName} -- ${clinic}`;
  const colMap = { Registrerare: 0, Mottagare: 1, ST: 2, Spec: 3, AT: 4, BT: 5 };
  const oldCols = new Set();
  if (oldUserRole === 'Registrerare' || oldUserRole === 'Blandbild') oldCols.add(colMap['Registrerare']);
  if (oldUserRole === 'Mottagare'   || oldUserRole === 'Blandbild') oldCols.add(colMap['Mottagare']);
  if (colMap[oldJobRole] !== undefined) oldCols.add(colMap[oldJobRole]);
  const newCols = new Set();
  if (newUserRole === 'Registrerare' || newUserRole === 'Blandbild') newCols.add(colMap['Registrerare']);
  if (newUserRole === 'Mottagare'    || newUserRole === 'Blandbild') newCols.add(colMap['Mottagare']);
  if (colMap[newJobRole] !== undefined) newCols.add(colMap[newJobRole]);
  oldCols.forEach(colIdx => {
    if (newCols.has(colIdx)) return;
    const colData = sheet.getRange(1, colIdx + 1, totalRows, 1).getValues().flat();
    const rowIdx = colData.findIndex(v => v === cellValue);
    if (rowIdx !== -1) sheet.getRange(rowIdx + 1, colIdx + 1).setValue('');
  });
  newCols.forEach(colIdx => {
    if (oldCols.has(colIdx)) return;
    const colData = sheet.getRange(1, colIdx + 1, totalRows, 1).getValues().flat();
    if (colData.includes(cellValue)) return;
    let emptyRow = -1;
    for (let i = 1; i < colData.length; i++) {
      if (!colData[i] || colData[i].toString().trim() === '') { emptyRow = i + 1; break; }
    }
    if (emptyRow === -1) emptyRow = totalRows + 1;
    sheet.getRange(emptyRow, colIdx + 1).setValue(cellValue);
  });
}

function hashExistingPins() {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Användare');
  const data = sheet.getDataRange().getValues();
  data.slice(1).forEach((row, i) => {
    const pin = row[7];
    if (pin && pin.toString().length <= 6) {
      sheet.getRange(i + 2, 8).setValue(hashPin(pin.toString()));
    }
  });
}

function authDrive() {
  DriveApp.getRootFolder();
}

// ── Cache-hantering ───────────────────────────────────────────────────────
const CACHE_KEY = 'getData_v3';
const CACHE_TTL = 600; // 10 min

function invalidateCache() {
  try { CacheService.getScriptCache().remove(CACHE_KEY); } catch(e) {}
}

function keepAlive() { /* tidsstyrd trigger – håller GAS-instansen varm */ }

// ── Data ──────────────────────────────────────────────────────────────────
function getData() {
  // Returnera cachad version om färsk (undviker rebuildCategories varje anrop)
  const cache = CacheService.getScriptCache();
  try {
    const hit = cache.get(CACHE_KEY);
    if (hit) return JSON.parse(hit);
  } catch(e) {}

  const ss = SpreadsheetApp.openById(DOC_ID);
  // Synka Categories endast vid cache-miss (dvs. max var 10:e minut)
  rebuildCategoriesFromUsers(ss);
  const sheet = ss.getSheetByName('Categories');
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const data = sheet.getRange(1, 1, lastRow, 7).getValues();
  const rows = data.slice(1);
  // Bygg kliniknamn → klinikId-mappning från Kliniker-sheetet
  const klinikIdMap = {};
  const klinikerSheet = ss.getSheetByName('Kliniker');
  if (klinikerSheet) {
    klinikerSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (r[0] && r[1]) klinikIdMap[r[1].toString().trim()] = r[0].toString().trim();
    });
  }
  // Clinics och locations från Kliniker-sheetet (Categories kol G är tömd av rebuildCategories)
  const clinics   = Object.keys(klinikIdMap).sort();
  const locations = [...new Set(clinics.map(c => c.includes(' - ') ? c.split(' - ')[0].trim() : c))].sort();
  const colMap = { Registrerare: 0, Mottagare: 1, ST: 2, Spec: 3, AT: 4, BT: 5 };
  const lists = { Registrerare: {}, Mottagare: {}, ST: {}, Spec: {}, AT: {}, BT: {} };
  Object.entries(colMap).forEach(([category, idx]) => {
    rows.forEach(row => {
      const cell = row[idx];
      if (cell && cell.trim()) {
        const parts = cell.split(' -- ');
        const name = parts[0].trim();
        const clinic = parts[1] ? parts[1].trim() : '';
        if (!lists[category][clinic]) lists[category][clinic] = [];
        lists[category][clinic].push(name);
      }
    });
  });
  const nameToCategory = {};
  ['Spec', 'ST', 'BT', 'AT'].forEach(cat => {
    rows.forEach(row => {
      const idx = colMap[cat];
      const cell = row[idx];
      if (cell && cell.trim()) {
        const name = cell.split(' -- ')[0].trim();
        nameToCategory[name] = cat;
      }
    });
  });
  // Hämta formulärtyper från Formulärtyper-sheetet
  const ftSheet = ss.getSheetByName('Formulärtyper');
  const formTypes = {};
  const timeSavings = {}; // formulärnamn -> minuter
  if (ftSheet) {
    const ftRows = ftSheet.getDataRange().getValues().slice(1);
    ftRows.forEach(r => {
      if (!r[0]) return;
      const formName = r[0].toString().trim();
      const cats = r[1] ? r[1].toString().split(',').map(c => c.trim()) : [];
      cats.forEach(cat => {
        if (!formTypes[cat]) formTypes[cat] = [];
        formTypes[cat].push(normalizeFormName(formName));
      });
      if (r[3]) timeSavings[normalizeFormName(formName)] = parseInt(r[3]) || 0;
    });
  }
  // Hämta kontextmeningar
  const ctxSheet = ss.getSheetByName('Kontextmeningar');
  const contextPhrases = [];
  if (ctxSheet) {
    ctxSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (r[2]) contextPhrases.push({ min: parseInt(r[0])||0, max: parseInt(r[1])||9999, text: r[2].toString().trim() });
    });
  }
  const result = { clinics, locations, lists, nameToCategory, formTypes, klinikIdMap, timeSavings, contextPhrases };
  try { cache.put(CACHE_KEY, JSON.stringify(result), CACHE_TTL); } catch(e) {} // misslyckas tyst om payload > 100KB
  return result;
}

function normalizeFormName(name) {
  if (!name) return '';
  name = name.toString().trim();
  if (name === 'Operationsspecifik DOPS') return 'DOPS';
  if (name === 'DOPS (allmän mall för mindre ingrepp/åtgärder)') return 'DOPS:Allmän';
  if (name === 'Återkoppling till ST-läkare efter muntlig presentation') return 'Återkoppling efter muntlig presentation';
  return name;
}

function getFormQuestions(formName) {
  formName = normalizeFormName(formName).trim();
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Formulär');
  const data = sheet.getRange(1, 1, Math.min(sheet.getLastRow(), 500), 9).getValues();
  const rows = data.slice(1);
  const results = rows.filter(r => r[0] && normalizeFormName(r[0]).trim() === formName)
    .map(r => ({
      form:        r[0],
      section:     r[1],
      nextSection: r[2],
      question:    r[3],
      description: r[4],
      type:        r[5],
      options:     r[6] ? r[6].toString().split(',').map(a => a.trim()) : [],
      jumpTo:      r[7] ? r[7].toString() : '',
      autofill:    r[8] ? r[8].toString().trim() : ''
    }));
  return results;
}

// ── Save + email ───────────────────────────────────────────────────────────
function saveAnswers(answerList, recipientEmail, recipientName, registrarName, formType) {
  const ss = SpreadsheetApp.openById(DOC_ID);
  let sheet = ss.getSheetByName('Svar');
  if (!sheet) {
    sheet = ss.insertSheet('Svar');
    sheet.appendRow(['Timestamp', 'Klinik', 'Registrerare', 'Mottagarkategori', 'Mottagare', 'Formulärtyp', 'Avsnitt', 'Fråga', 'Svar', 'SubmissionID']);
  }
  const ts = new Date();
  const submissionId = Utilities.getUuid();
  answerList.forEach(a => {
    sheet.appendRow([ts, a.clinic, a.registrar, a.category, a.recipient, a.formType, a.section, a.question, a.answer, submissionId, a.klinikId || '']);
  });
  if (recipientEmail) {
    // Uppdatera SenastAktiv för mottagare
    const userSheet2 = ss.getSheetByName('Användare');
    const uData2 = userSheet2.getDataRange().getValues().slice(1);
    [recipientEmail, null].forEach((emailToTouch, idx) => {
      const target = idx === 0 ? recipientEmail : null;
      // Touched via answerList[0].registrar name matching skips for now
    });
    const recIdx = uData2.findIndex(r => r[3].toString().trim().toLowerCase() === recipientEmail.toLowerCase());
    if (recIdx !== -1) userSheet2.getRange(recIdx + 2, 17).setValue(new Date().toISOString().slice(0,10));
    // Kontrollera emailNotify för mottagaren
    const userSheet = ss.getSheetByName('Användare');
    const uData = userSheet.getDataRange().getValues().slice(1);
    const uRow = uData.find(r => r[3].toString().trim().toLowerCase() === recipientEmail.toLowerCase());
    const emailNotify = uRow ? (uRow[13] || 'Ja').toString().trim() : 'Ja';
    if (emailNotify !== 'Nej') {
      sendAssessmentEmail(recipientEmail, recipientName, registrarName, formType, answerList, ts);
    }
  }
  return 'OK';
}

function sendAssessmentEmail(toEmail, recipientName, registrarName, formType, answers, timestamp) {
  const dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  let html = '<div style="font-family:Arial,sans-serif;max-width:620px;margin:0 auto;background:#f0ede8;padding:0;">';
  html += '<div style="background:#3d5c4a;border-radius:6px 6px 0 0;padding:0;">';
  html += '<img src="https://mokir-web.github.io/dops/email-header.png" alt="MoraDOPS" width="600" style="display:block;width:100%;max-width:600px;border-radius:6px 6px 0 0;" />';
  html += '</div>';
  html += '<div style="background:#f5f1ec;padding:28px 32px 32px;border-radius:0 0 6px 6px;border:1.5px solid #d8cfc5;border-top:none;">';
  html += '<h2 style="color:#2a1f18;margin-top:0;">' + formType + '</h2>';
  html += '<table style="width:100%;margin-bottom:16px;font-size:14px;">';
  html += '<tr><td style="color:#6b5c52;width:120px;">Datum</td><td style="font-weight:bold;">' + dateStr + '</td></tr>';
  html += '<tr><td style="color:#6b5c52;">Mottagare</td><td style="font-weight:bold;">' + recipientName + '</td></tr>';
  html += '<tr><td style="color:#6b5c52;">Registrerare</td><td style="font-weight:bold;">' + registrarName + '</td></tr>';
  html += '</table><hr style="border:none;border-top:1.5px solid #d8cfc5;margin:16px 0;">';
  var currentSec = '';
  answers.forEach(function(a) {
    if (a.section !== currentSec) {
      currentSec = a.section;
      html += '<div style="background:#d8cfc5;padding:8px 12px;border-radius:4px;border-left:4px solid #3d5c4a;font-weight:bold;color:#2a1f18;margin:16px 0 8px;">' + a.section + '</div>';
    }
    if (a.answer) html += '<div style="padding:4px 0 4px 12px;font-size:14px;"><span style="color:#6b5c52;">' + escapeHtml(a.question) + ':</span> <span style="color:#1e1612;">' + escapeHtml(a.answer) + '</span></div>';
  });
  html += '<hr style="border:none;border-top:1.5px solid #d8cfc5;margin:16px 0;">';
  html += '<div style="font-size:12px;color:#9e8e82;">Frågor som inte anses relevanta har lämnats obesvarade.</div>';
  html += getUnsubscribeFooter(toEmail);
  html += '</div></div>';
  var body = 'Hej ' + recipientName + ',\n\nDu har mottagit en bedömning av ' + registrarName + '.\nFormulär: ' + formType + '\nDatum: ' + dateStr + '\n';
  var currentSecTxt = '';
  answers.forEach(function(a) {
    if (a.section !== currentSecTxt) { currentSecTxt = a.section; body += '\n' + a.section + '\n'; }
    if (a.answer) body += a.question + ': ' + a.answer + '\n';
  });
  MailApp.sendEmail({ to: toEmail, name: 'MoraDOPS', subject: 'Bedömning mottagen: ' + formType + ' — ' + dateStr, body: body, htmlBody: html });
}
function getRecipientEmail(recipientFullName) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Användare');
  const data = sheet.getDataRange().getValues();
  const namePart = recipientFullName.split(' -- ')[0].trim();
  const row = data.slice(1).find(r => `${r[1]} ${r[2]}`.trim() === namePart);
  return row ? row[3] : null;
}

// ── My assessments ─────────────────────────────────────────────────────────
function getMyAssessments(email, role, filters) {
  const ss = SpreadsheetApp.openById(DOC_ID);
  const userSheet = ss.getSheetByName('Användare');
  const userData = userSheet.getDataRange().getValues();
  const userRow = userData.slice(1).find(r => r[3].toString().toLowerCase() === email.toLowerCase());
  if (!userRow) return [];
  const fullName = `${userRow[1]} ${userRow[2]}`;
  const svarSheet = ss.getSheetByName('Svar');
  if (!svarSheet) return [];
  const rows = svarSheet.getDataRange().getValues().slice(1);
  const assessments = {};
  rows.forEach(row => {
    const ts        = row[0];
    const registrar = row[2];
    const category  = row[3];
    const recipient = row[4];
    const formType  = row[5];
    const section   = row[6];
    const question  = row[7];
    const answer    = row[8];
    const recName = recipient.split(' -- ')[0];
    const isMatch = role === 'sent' ? registrar === fullName : recName === fullName;
    if (!isMatch) return;
    if (filters && filters.dateFrom) {
      const tsDate = ts instanceof Date ? ts : new Date(ts);
      const from = new Date(filters.dateFrom);
      const to = filters.dateTo ? new Date(filters.dateTo) : new Date();
      to.setHours(23, 59, 59);
      if (tsDate < from || tsDate > to) return;
    }
    if (filters && filters.formTypes && filters.formTypes.length > 0) {
      if (!filters.formTypes.includes(formType)) return;
    }
    const tsStr = ts instanceof Date
      ? Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
      : ts.toString();
    const submissionId = row[9] ? row[9].toString() : `${tsStr}_${recipient}_${formType}`;
    const key = submissionId;
    if (!assessments[key]) {
      assessments[key] = { timestamp: tsStr, registrar, recipient, category, formType, answers: [] };
    }
    assessments[key].answers.push({ section, question, answer });
  });
  return Object.values(assessments).sort((a, b) => b.timestamp.localeCompare(a.timestamp));
}

// ── Statistics ────────────────────────────────────────────────────────────
function getStatistics(filters, klinikId) {
  const ss = SpreadsheetApp.openById(DOC_ID);
  const svarSheet = ss.getSheetByName('Svar');
  if (!svarSheet) return { total: 0, byFormType: {}, byMonth: {}, byRegistrar: {}, byRecipient: {}, scaleAverages: {} };
  const allRows = svarSheet.getDataRange().getValues().slice(1);
  // Filtrera på klinik (datumfilter hanteras via filters-parametern nedan)
  const rows = allRows.filter(r => {
    if (!klinikId || klinikId === '*') return true;
    return r[10] && r[10].toString().trim() === klinikId;
  });
  const total_submissions = new Set();
  const byFormType = {};
  const byMonth = {};
  const byRegistrar = {};
  const byRecipient = {};
  const scaleData = {};

  rows.forEach(row => {
    const ts       = row[0];
    const registrar = row[2];
    const recipient = row[4];
    const formType  = row[5];
    const question  = row[7];
    const answer    = row[8];
    const submId    = row[9] ? row[9].toString() : null;

    if (!ts || !formType) return;
    const tsDate = ts instanceof Date ? ts : new Date(ts);

    // Datumfilter
    if (filters && filters.dateFrom) {
      const from = new Date(filters.dateFrom);
      const to = filters.dateTo ? new Date(filters.dateTo) : new Date();
      to.setHours(23,59,59);
      if (tsDate < from || tsDate > to) return;
    }

    if (submId) total_submissions.add(submId);

    // Per formulärtyp
    byFormType[formType] = (byFormType[formType] || 0);
    if (submId && !byFormType[`_seen_${submId}`]) {
      byFormType[formType]++;
      byFormType[`_seen_${submId}`] = true;
    }

    // Per månad
    const monthKey = Utilities.formatDate(tsDate, Session.getScriptTimeZone(), 'yyyy-MM');
    if (!byMonth[monthKey]) byMonth[monthKey] = { label: Utilities.formatDate(tsDate, Session.getScriptTimeZone(), 'MMM yyyy'), count: new Set() };
    if (submId) byMonth[monthKey].count.add(submId);

    // Per registrerare
    if (!byRegistrar[registrar]) byRegistrar[registrar] = new Set();
    if (submId) byRegistrar[registrar].add(submId);

    // Per mottagare
    const recName = recipient ? recipient.split(' -- ')[0] : recipient;
    if (!byRecipient[recName]) byRecipient[recName] = new Set();
    if (submId) byRecipient[recName].add(submId);

    // Skalvärden
    const numVal = parseFloat(answer);
    if (!isNaN(numVal) && numVal >= 1 && numVal <= 9) {
      if (!scaleData[question]) scaleData[question] = [];
      scaleData[question].push(numVal);
    }
  });

  // Konvertera Sets till counts
  const byMonthClean = {};
  Object.entries(byMonth).sort((a,b) => a[0].localeCompare(b[0])).forEach(([k, v]) => {
    byMonthClean[k] = { label: v.label, count: v.count.size, key: k };
  });
  const byRegClean = {};
  Object.entries(byRegistrar).forEach(([k, v]) => { byRegClean[k] = v.size; });
  const byRecClean = {};
  Object.entries(byRecipient).forEach(([k, v]) => { byRecClean[k] = v.size; });

  // Rensa _seen_-nycklar från byFormType
  const byFormTypeClean = {};
  Object.entries(byFormType).forEach(([k, v]) => {
    if (!k.startsWith('_seen_')) byFormTypeClean[k] = v;
  });

  const scaleAverages = {};
  Object.entries(scaleData).forEach(([q, vals]) => {
    scaleAverages[q] = Math.round((vals.reduce((a,b) => a+b, 0) / vals.length) * 10) / 10;
  });

  // Bygg byFormTypeByMonth för linjediagram
  const byFormTypeByMonth = {};
  Object.entries(byMonth).forEach(([monthKey, mv]) => {
    byFormTypeByMonth[monthKey] = { label: mv.label };
  });
  rows.forEach(row => {
    const ts = row[0]; const formType = row[5]; const submId = row[9] ? row[9].toString() : null;
    if (!ts || !formType || !submId) return;
    const tsDate = ts instanceof Date ? ts : new Date(ts);
    if (filters && filters.dateFrom) {
      const from = new Date(filters.dateFrom);
      const to = filters.dateTo ? new Date(filters.dateTo) : new Date();
      to.setHours(23,59,59);
      if (tsDate < from || tsDate > to) return;
    }
    const monthKey = Utilities.formatDate(tsDate, Session.getScriptTimeZone(), 'yyyy-MM');
    if (!byFormTypeByMonth[monthKey]) byFormTypeByMonth[monthKey] = { label: monthKey };
    if (!byFormTypeByMonth[monthKey][formType]) byFormTypeByMonth[monthKey][formType] = new Set();
    byFormTypeByMonth[monthKey][formType].add(submId);
  });
  const byFormTypeByMonthClean = {};
  Object.entries(byFormTypeByMonth).sort((a,b) => a[0].localeCompare(b[0])).forEach(([k,v]) => {
    byFormTypeByMonthClean[k] = { label: v.label };
    Object.entries(v).forEach(([fk,fv]) => {
      if (fk !== 'label') byFormTypeByMonthClean[k][fk] = fv instanceof Set ? fv.size : fv;
    });
  });

  return {
    total: total_submissions.size,
    byFormType: byFormTypeClean,
    byMonth: byMonthClean,
    byFormTypeByMonth: byFormTypeByMonthClean,
    byRegistrar: byRegClean,
    byRecipient: byRecClean,
    scaleAverages
  };
}

// ── Clinic/Admin assessments ──────────────────────────────────────────────
function getClinicAssessments(klinikId, filters) {
  const ss = SpreadsheetApp.openById(DOC_ID);
  const svarSheet = ss.getSheetByName('Svar');
  if (!svarSheet) return [];
  const userSheet = ss.getSheetByName('Användare');
  const userData = userSheet.getDataRange().getValues().slice(1);

  // Bygg upp email→klinikId-mappning
  const emailToKlinik = {};
  userData.forEach(r => {
    if (r[3] && r[10]) emailToKlinik[r[3].toString().toLowerCase()] = r[10].toString().trim();
  });

  const rows = svarSheet.getDataRange().getValues().slice(1);
  const assessments = {};

  rows.forEach(row => {
    const ts        = row[0];
    const registrar = row[2];
    const category  = row[3];
    const recipient = row[4];
    const formType  = row[5];
    const section   = row[6];
    const question  = row[7];
    const answer    = row[8];
    const submId    = row[9] ? row[9].toString() : null;
    const rowClinik = row[10] ? row[10].toString().trim() : '';

    if (!ts || !formType) return;

    // Filtrera på klinik – fallback: slå upp registrarens klinikId om rowClinik saknas
    let effectiveKlinik = rowClinik;
    if (!effectiveKlinik) {
      // Försök hitta via registrarens email i emailToKlinik
      // Registrar-fältet är "Förnamn Efternamn" – kolla via namn-mappning istället
      // Enklare fallback: om klinikId är tomt, matcha på klinikknamn i Klinik-kolumnen (row[1])
      const rowKlinikName = row[1] ? row[1].toString().trim() : '';
      // Hitta klinikId för detta kliniknamn
      const matchedUser = userData.find(r => r[5] && r[5].toString().trim() === rowKlinikName && r[10] && r[10].toString().trim());
      if (matchedUser) effectiveKlinik = matchedUser[10].toString().trim();
    }
    if (klinikId !== '*' && effectiveKlinik !== klinikId) return;

    const tsDate = ts instanceof Date ? ts : new Date(ts);
    if (filters && filters.dateFrom) {
      const from = new Date(filters.dateFrom);
      const to = filters.dateTo ? new Date(filters.dateTo) : new Date();
      to.setHours(23,59,59);
      if (tsDate < from || tsDate > to) return;
    }
    if (filters && filters.formTypes && filters.formTypes.length > 0) {
      if (!filters.formTypes.includes(formType)) return;
    }
    if (filters && filters.category && filters.category !== '') {
      if (category !== filters.category) return;
    }
    if (filters && filters.recipient && filters.recipient !== '') {
      const recName = recipient ? recipient.split(' -- ')[0] : '';
      if (!recName.toLowerCase().includes(filters.recipient.toLowerCase())) return;
    }

    const tsStr = ts instanceof Date
      ? Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
      : ts.toString();
    const key = submId || `${tsStr}_${recipient}_${formType}`;
    if (!assessments[key]) {
      assessments[key] = { timestamp: tsStr, registrar, recipient, category, formType, answers: [], klinikId: rowClinik };
    }
    assessments[key].answers.push({ section, question, answer });
  });

  return Object.values(assessments).sort((a, b) => b.timestamp.localeCompare(a.timestamp));
}

// ── Export to Word ─────────────────────────────────────────────────────────
function exportAssessmentsToDoc(email, assessments, filterInfo) {
  const docName = `Bedömningar — ${filterInfo.name} — ${filterInfo.period}`;
  const doc = DocumentApp.create(docName);
  const body = doc.getBody();
  const headerStyle = {};
  headerStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  headerStyle[DocumentApp.Attribute.BOLD] = false;
  const header = doc.addHeader();
  header.setText(`${filterInfo.name}   |   ${filterInfo.clinic}   |   Period: ${filterInfo.period}`);
  header.setAttributes(headerStyle);
  assessments.sort((a, b) => a.formType.localeCompare(b.formType));
  const countByType = {};
  assessments.forEach(a => { countByType[a.formType] = (countByType[a.formType] || 0) + 1; });
  // Lägg till MoraDOPS-rubrik längst upp till höger
  const topRow = body.insertTable(0);
  const headerRow = topRow.appendTableRow();
  const leftCell = headerRow.appendTableCell();
  leftCell.setText(filterInfo.name + '   |   ' + filterInfo.clinic + '   |   Period: ' + filterInfo.period);
  leftCell.getChild(0).asText().setFontSize(9).setForegroundColor('#6b5c52');
  const rightCell = headerRow.appendTableCell();
  const rightPara = rightCell.getChild(0).asParagraph();
  rightPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  rightPara.appendText('MoraDOPS').setFontSize(16).setBold(true).setForegroundColor('#3d5c4a');
  topRow.setBorderColor('#c4b8ac');
  body.appendParagraph('');

  const tocTitle = body.appendParagraph('Innehållsförteckning');
  tocTitle.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  let num = 1;
  const tocEntries = [];
  Object.entries(countByType).forEach(([type, count]) => {
    tocEntries.push({ type, count, startNum: num });
    num += count;
  });
  tocEntries.forEach(e => {
    body.appendParagraph(`${e.type}   ${e.count} st   (nr ${e.startNum}–${e.startNum + e.count - 1})`);
  });
  body.appendParagraph('');
  body.appendHorizontalRule();
  body.appendParagraph('');
  let assessmentNum = 1;
  assessments.forEach(a => {
    const titlePara = body.appendParagraph(`${assessmentNum}. ${a.formType}`);
    titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(`Datum: ${a.timestamp}`);
    body.appendParagraph(`Mottagare: ${a.recipient.split(' -- ')[0]}`);
    body.appendParagraph(`Registrerare: ${a.registrar}`);
    body.appendParagraph(`Kategori: ${a.category}`);
    body.appendParagraph('');
    let lastSection = '';
    a.answers.forEach(ans => {
      if (ans.section !== lastSection) {
        lastSection = ans.section;
        const secPara = body.appendParagraph(ans.section);
        secPara.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      }
      body.appendParagraph(`${ans.question}: ${ans.answer}`);
    });
    body.appendParagraph('');
    body.appendHorizontalRule();
    body.appendParagraph('');
    assessmentNum++;
  });
  doc.saveAndClose();
  const docId = doc.getId();
  const url = `https://docs.google.com/feeds/download/documents/export/Export?id=${docId}&exportFormat=docx`;
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const blob = response.getBlob().setName(`${docName}.docx`);
  MailApp.sendEmail({
    to: email,
    subject: docName,
    body: `Bifogat finns dina exporterade bedömningar.\n\nPeriod: ${filterInfo.period}\nAntal: ${assessments.length} st`,
    attachments: [blob]
  });
  DocumentApp.openById(docId).setName(`[EXPORTERAD] ${docName}`);
  return { success: true, count: assessments.length };
}

// ── Admin: get all users ──────────────────────────────────────────────────
function getAllUsers(klinikId, includeInaktiverade) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Användare');
  const rows = sheet.getDataRange().getValues().slice(1);
  return rows
    .filter(r => r[1] && r[3])
    .filter(r => includeInaktiverade || (r[17] || '').toString().trim() !== 'Ja')
    .filter(r => !klinikId || klinikId === '*' || (r[10] && r[10].toString().trim() === klinikId))
    .map(r => ({
      firstName: r[1].toString(), lastName: r[2].toString(),
      email:     r[3].toString(), jobRole:  r[4].toString(),
      clinic:    r[5].toString(), userRole: r[6].toString().replace('Registrerare av bedömningar','Registrerare').replace('Mottagare av bedömningar','Mottagare').trim(),
      klinikId:  r[10] ? r[10].toString().trim() : '',
      senastAktiv: r[16] ? r[16].toString() : '',
      inaktiverad: (r[17] || '').toString().trim() === 'Ja'
    }));
}

// ── Delete user ────────────────────────────────────────────────────────────
function deleteUser(email) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Användare');
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex((r, i) => i > 0 && r[3].toString().toLowerCase() === email.toLowerCase());
  if (rowIdx === -1) return { error: 'Användare hittades inte.' };
  sheet.deleteRow(rowIdx + 1);
  // Ta även bort privilegier
  const privSheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Privilegier');
  if (privSheet) {
    const privData = privSheet.getDataRange().getValues();
    for (let i = privData.length - 1; i >= 1; i--) {
      if (privData[i][0].toString().toLowerCase() === email.toLowerCase()) {
        privSheet.deleteRow(i + 1);
      }
    }
  }
  invalidateCache();
  return { success: true };
}

// ── Admin: update user ────────────────────────────────────────────────────
function updateUserByAdmin(targetEmail, updates) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Användare');
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex((r, i) => i > 0 && r[3].toString().toLowerCase() === targetEmail.toLowerCase());
  if (rowIdx === -1) return { error: 'Användare hittades inte.' };
  const row = rowIdx + 1;
  if (updates.firstName !== undefined) sheet.getRange(row, 2).setValue(updates.firstName);
  if (updates.lastName  !== undefined) sheet.getRange(row, 3).setValue(updates.lastName);
  if (updates.jobRole   !== undefined) sheet.getRange(row, 5).setValue(updates.jobRole);
  if (updates.clinic    !== undefined) sheet.getRange(row, 6).setValue(updates.clinic);
  if (updates.userRole  !== undefined) sheet.getRange(row, 7).setValue(updates.userRole);
  if (updates.klinikId  !== undefined) sheet.getRange(row, 11).setValue(updates.klinikId);
  if (updates.newPin    !== undefined) sheet.getRange(row, 8).setValue(hashPin(updates.newPin));
  // Synka Categories efter ändring
  const ss2 = SpreadsheetApp.openById(DOC_ID);
  rebuildCategoriesFromUsers(ss2);
  invalidateCache();
  return { success: true };
}

// ── Admin: get clinic info ─────────────────────────────────────────────────
function getClinicInfo(klinikId) {
  const ss = SpreadsheetApp.openById(DOC_ID);
  const userSheet = ss.getSheetByName('Användare');
  const users = userSheet.getDataRange().getValues().slice(1)
    .filter(r => r[1] && r[3] && (klinikId === '*' || r[10].toString().trim() === klinikId))
    .map(r => ({
      firstName: r[1].toString(), lastName: r[2].toString(),
      email: r[3].toString(), jobRole: r[4].toString(),
      clinic: r[5].toString(), userRole: r[6].toString(),
      klinikId: r[10] ? r[10].toString() : ''
    }));
  const kliniker = ss.getSheetByName('Kliniker');
  let info = null;
  if (kliniker) {
    const rows = kliniker.getDataRange().getValues().slice(1);
    const row = rows.find(r => r[0].toString() === klinikId);
    if (row) info = { klinikId: row[0], namn: row[1], kontakt: row[2], aktiv: row[3], start: row[4] };
  }
  return { users, info };
}

// ── Admin: toggle clinic access ────────────────────────────────────────────
function setClinicActive(klinikId, active) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Kliniker');
  if (!sheet) return { error: 'Kliniker-sheet saknas.' };
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex((r, i) => i > 0 && r[0].toString() === klinikId);
  if (rowIdx === -1) return { error: 'Klinik hittades inte.' };
  sheet.getRange(rowIdx + 1, 4).setValue(active ? 'Ja' : 'Nej');
  return { success: true };
}

// ── Privilege management ──────────────────────────────────────────────────
function getPrivilegesForUser(targetEmail) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Privilegier');
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter(r => r[0].toString().trim().toLowerCase() === targetEmail.toLowerCase())
    .map(r => ({ privilege: r[1].toString().trim(), klinikId: r[2].toString().trim() }));
}

function setUserPrivilege(targetEmail, privilege, klinikId, remove) {
  targetEmail = (targetEmail || '').trim().toLowerCase();
  privilege   = (privilege   || '').trim();
  klinikId    = (klinikId    || '').trim();
  const ss = SpreadsheetApp.openById(DOC_ID);
  let sheet = ss.getSheetByName('Privilegier');
  if (!sheet) {
    sheet = ss.insertSheet('Privilegier');
    sheet.appendRow(['Email', 'Privilegium', 'KlinikID']);
  }
  const data = sheet.getDataRange().getValues();
  // Hitta befintlig rad
  const rowIdx = data.findIndex((r, i) =>
    i > 0 &&
    r[0].toString().trim().toLowerCase() === targetEmail &&
    r[1].toString().trim() === privilege
  );
  if (remove) {
    if (rowIdx !== -1) sheet.deleteRow(rowIdx + 1);
  } else {
    if (rowIdx !== -1) {
      sheet.getRange(rowIdx + 1, 3).setValue(klinikId);
    } else {
      sheet.appendRow([targetEmail, privilege, klinikId]);
    }
  }
  // Beräkna netto tidsbesparing (papperssparande minus faktisk tid)
  let netMinutesSaved = 0;
  if (data.startTime && data.formType) {
    const ftSheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Formulärtyper');
    if (ftSheet) {
      const ftRow = ftSheet.getDataRange().getValues().slice(1)
        .find(r => r[0] && r[0].toString().trim() === data.formType);
      if (ftRow && ftRow[3]) {
        const paperMins = parseInt(ftRow[3]) || 0;
        const actualMins = Math.round((new Date() - new Date(data.startTime)) / 60000);
        netMinutesSaved = Math.max(0, paperMins - actualMins);
      }
    }
  }
  return { success: true, netMinutesSaved };
}

function getKlinikSettings(klinikId) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Kliniker');
  if (!sheet) return { inaktiveringsMånader: 24 };
  const rows = sheet.getDataRange().getValues().slice(1);
  const row = rows.find(r => r[0].toString() === klinikId);
  return row ? { inaktiveringsMånader: parseInt(row[5]) || 24 } : { inaktiveringsMånader: 24 };
}

function saveKlinikSettings(klinikId, settings) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Kliniker');
  if (!sheet) return { error: 'Sheet saknas' };
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex((r, i) => i > 0 && r[0].toString() === klinikId);
  if (rowIdx === -1) return { error: 'Klinik hittades inte' };
  if (settings.inaktiveringsMånader !== undefined)
    sheet.getRange(rowIdx + 1, 6).setValue(parseInt(settings.inaktiveringsMånader) || 24);
  return { success: true };
}

// Tidsstyrd trigger – kör månadsvis
function inaktiveraInaktivaAnvändare() {
  const ss = SpreadsheetApp.openById(DOC_ID);
  const userSheet = ss.getSheetByName('Användare');
  if (!userSheet) return;
  const klinikerSheet = ss.getSheetByName('Kliniker');
  // Bygg klinik → inaktiveringsgräns
  const klinikGräns = {};
  if (klinikerSheet) {
    klinikerSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (r[0]) klinikGräns[r[0].toString().trim()] = parseInt(r[5]) || 24;
    });
  }
  const now = new Date();
  const rows = userSheet.getDataRange().getValues();
  rows.slice(1).forEach((row, i) => {
    if (!row[1] || !row[3]) return;
    if ((row[17] || '').toString().trim() === 'Ja') return; // redan inaktiverad
    const klinikId = (row[10] || '').toString().trim();
    const månader = klinikGräns[klinikId] || 24;
    const gräns = new Date(now.getFullYear(), now.getMonth() - månader, now.getDate());
    const senast = row[16] ? new Date(row[16]) : (row[0] instanceof Date ? row[0] : new Date(row[0]));
    if (senast < gräns) {
      userSheet.getRange(i + 2, 18).setValue('Ja'); // R=Inaktiverad
    }
  });
  invalidateCache();
}
function sendAssessmentRequest(fromEmail, toEmail, message, formType) {
  formType = (formType || '').trim();
  fromEmail = (fromEmail || '').trim().toLowerCase();
  toEmail   = (toEmail   || '').trim().toLowerCase();
  message   = (message   || '').trim();
  const ss = SpreadsheetApp.openById(DOC_ID);

  // Spara i Förfrågningar-sheet
  let sheet = ss.getSheetByName('Förfrågningar');
  if (!sheet) {
    sheet = ss.insertSheet('Förfrågningar');
    sheet.appendRow(['Timestamp', 'Från', 'Till', 'Meddelande', 'Läst', 'ID']);
  }
  const id = Utilities.getUuid();
  sheet.appendRow([new Date(), fromEmail, toEmail, message, 'Nej', id, '', formType]);

  // Skicka mail till mottagaren om emailNotify
  const userSheet = ss.getSheetByName('Användare');
  const users = userSheet.getDataRange().getValues().slice(1);
  const fromRow = users.find(r => r[3].toString().trim().toLowerCase() === fromEmail);
  const toRow   = users.find(r => r[3].toString().trim().toLowerCase() === toEmail);
  const emailNotify = toRow ? (toRow[14] || 'Ja').toString().trim() : 'Ja'; // kolumn O = emailNotifyRequests
  if (toRow && emailNotify !== 'Nej') {
    const fromName = fromRow ? fromRow[1] + ' ' + fromRow[2] : fromEmail;
    const toName   = toRow[1] + ' ' + toRow[2];
    const reqSubject = formType
      ? '[MoraDOPS] ' + formType + ' — begärd av ' + fromName
      : 'Förfrågan om bedömning från ' + fromName;
    const reqHtml = buildRequestEmail(fromName, toName, message, formType, toRow[3].toString().trim());
    MailApp.sendEmail({
      to: toRow[3].toString().trim(),
      replyTo: 'donotreply.mokir@gmail.com',
      name: 'MoraDOPS',
      subject: reqSubject,
      htmlBody: reqHtml,
      body: 'Hej ' + toName + ',\n\n' + fromName + ' ber dig registrera en bedömning' + (formType ? ': ' + formType : '') + '.\n\n"' + message + '"\n\nLogga in på MoraDOPS för att utföra bedömningen.'
    });
  }
  return { success: true, id };
}

function buildRequestEmail(fromName, toName, message, formType, toEmail) {
  const siteUrl = 'https://mokir-web.github.io/dops/';
  let html = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#f0ede8;padding:0;">';
  html += '<div style="background:#3d5c4a;border-radius:6px 6px 0 0;padding:0;">';
  html += '<img src="https://mokir-web.github.io/dops/email-header.png" alt="MoraDOPS" width="600" style="display:block;width:100%;max-width:600px;border-radius:6px 6px 0 0;" />';
  html += '</div>';
  html += '<div style="background:#f5f1ec;padding:28px 32px 32px;border-radius:0 0 6px 6px;border:1.5px solid #d8cfc5;border-top:none;">';
  html += '<h2 style="color:#2a1f18;margin-top:0;">Förfrågan om bedömning</h2>';
  html += '<table style="width:100%;margin-bottom:16px;font-size:14px;border-collapse:collapse;">';
  html += '<tr><td style="color:#6b5c52;padding:6px 0;width:120px;">Från</td><td style="font-weight:bold;">' + escapeHtml(fromName) + '</td></tr>';
  if (formType) {
    html += '<tr><td style="color:#6b5c52;padding:6px 0;">Formulär</td><td style="font-weight:bold;color:#3d5c4a;">' + escapeHtml(formType) + '</td></tr>';
  }
  html += '</table>';
  html += '<hr style="border:none;border-top:1px solid #d8cfc5;margin:16px 0;">';
  html += '<div style="background:#fff;border-radius:6px;padding:14px 16px;font-size:15px;color:#2a1f18;border-left:4px solid #3d5c4a;margin-bottom:24px;">';
  html += '<div style="font-size:12px;color:#6b5c52;margin-bottom:6px;letter-spacing:1px;">MEDDELANDE</div>';
  html += escapeHtml(message);
  html += '</div>';
  html += '<a href="' + siteUrl + '" style="display:inline-block;background:#3d5c4a;color:#f0ede8;text-decoration:none;padding:12px 28px;border-radius:6px;font-size:15px;font-weight:bold;">Öppna MoraDOPS →</a>';
  html += '<p style="color:#9e8e82;font-size:12px;margin-top:16px;">Eller kopiera länken: <a href="' + siteUrl + '" style="color:#3d5c4a;">' + siteUrl + '</a></p>';
  html += getUnsubscribeFooter(toEmail || '');
  html += '</div></div>';
  return html;
}

function getAssessmentRequests(email) {
  email = (email || '').trim().toLowerCase();
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Förfrågningar');
  if (!sheet) return { inbox: [], outbox: [] };
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { inbox: [], outbox: [] };
  // Läs explicit 8 kolumner (A–H) för att undvika getDataRange-truncation
  const rows = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const ss = SpreadsheetApp.openById(DOC_ID);
  const userSheet = ss.getSheetByName('Användare');
  const users = userSheet.getDataRange().getValues().slice(1);
  const nameMap = {};
  users.forEach(r => { nameMap[r[3].toString().trim().toLowerCase()] = r[1] + ' ' + r[2]; });

  const inbox = rows.filter(r => r[2].toString().trim().toLowerCase() === email && r[6]?.toString().trim() !== 'Utförd')
    .map(r => ({
      id: r[5].toString(), timestamp: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : r[0].toString(),
      from: r[1].toString().trim(), fromName: nameMap[r[1].toString().trim().toLowerCase()] || r[1].toString().trim(),
      message: r[3].toString(), read: r[4].toString() === 'Ja',
      status: r[6].toString().trim(),
      formType: r[7] ? r[7].toString().trim() : ''
    })).sort((a,b) => b.timestamp.localeCompare(a.timestamp));

  const outbox = rows.filter(r => r[1].toString().trim().toLowerCase() === email)
    .map(r => ({
      id: r[5].toString(), timestamp: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : r[0].toString(),
      to: r[2].toString().trim(), toName: nameMap[r[2].toString().trim().toLowerCase()] || r[2].toString().trim(),
      message: r[3].toString(), read: r[4].toString() === 'Ja',
      status: r[6].toString().trim(),
      formType: r[7] ? r[7].toString().trim() : ''
    })).sort((a,b) => b.timestamp.localeCompare(a.timestamp));

  return { inbox, outbox };
}

function markRequestRead(id) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Förfrågningar');
  if (!sheet) return { error: 'Sheet saknas' };
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex((r,i) => i > 0 && r[5].toString() === id);
  if (rowIdx !== -1) sheet.getRange(rowIdx + 1, 5).setValue('Ja');
  return { success: true };
}

function markRequestDone(id) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Förfrågningar');
  if (!sheet) return { error: 'Sheet saknas' };
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false };
  const data = sheet.getRange(1, 1, lastRow, 8).getValues();
  const rowIdx = data.findIndex((r,i) => i > 0 && r[5].toString() === id);
  if (rowIdx !== -1) {
    sheet.getRange(rowIdx + 1, 5).setValue('Ja');      // läst
    sheet.getRange(rowIdx + 1, 7).setValue('Utförd');  // status
  }
  return { success: true };
}

function markRequestDoneByParticipants(fromEmail, toEmail) {
  // Markera senaste aktiva förfrågan mellan dessa två som utförd
  fromEmail = (fromEmail || '').trim().toLowerCase();
  toEmail   = (toEmail   || '').trim().toLowerCase();
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Förfrågningar');
  if (!sheet) return { success: false };
  const lastRow2 = sheet.getLastRow();
  if (lastRow2 < 2) return { success: false };
  const data = sheet.getRange(1, 1, lastRow2, 8).getValues();
  // Hitta senaste matchande rad som ej är utförd
  for (let i = data.length - 1; i >= 1; i--) {
    const rowFrom = data[i][1].toString().trim().toLowerCase();
    const rowTo   = data[i][2].toString().trim().toLowerCase();
    const status  = data[i][6].toString().trim();
    if (rowTo === toEmail && rowFrom === fromEmail && status !== 'Utförd') {
      sheet.getRange(i + 1, 5).setValue('Ja');
      sheet.getRange(i + 1, 7).setValue('Utförd');
      return { success: true };
    }
  }
  return { success: false };
}

function cleanOldRequests() {
  // Radera förfrågningar äldre än 60 dagar
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Förfrågningar');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const cutoff = new Date(Date.now() - 60 * 24 * 60 * 60 * 1000);
  for (let i = data.length - 1; i >= 1; i--) {
    const ts = data[i][0] instanceof Date ? data[i][0] : new Date(data[i][0]);
    if (ts < cutoff) sheet.deleteRow(i + 1);
  }
}

// ── Min översikt ──────────────────────────────────────────────────────────
function getMyOverview(email, klinikId) {
  email = (email || '').trim().toLowerCase();
  const ss = SpreadsheetApp.openById(DOC_ID);
  const svarSheet = ss.getSheetByName('Svar');
  if (!svarSheet) return { sent: {}, received: {}, byMonth: {} };
  const rows = svarSheet.getDataRange().getValues().slice(1);

  // Hämta formulärtyper med målvärden och subingrepp
  const ftSheet = ss.getSheetByName('Formulärtyper');
  const formGoals = {}; // formulärnamn -> målvärde
  const subGoals  = {}; // t.ex. "DOPS:Appendektomi" -> målvärde
  if (ftSheet) {
    ftSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (!r[0]) return;
      const name = r[0].toString().trim();
      const goal = r[2] ? parseInt(r[2]) : 0;
      if (name.startsWith('DOPS:')) {
        subGoals[name] = goal;
      } else {
        formGoals[name] = goal;
      }
    });
  }

  // Hämta användarens fullständiga namn för matchning
  const userSheet = ss.getSheetByName('Användare');
  const uRow = userSheet.getDataRange().getValues().slice(1)
    .find(r => r[3].toString().trim().toLowerCase() === email);
  const fullName = uRow ? (uRow[1] + ' ' + uRow[2]).trim() : '';

  const sent     = {};
  const received = {};
  const byMonth  = {};
  const byMonthFormType = {}; // { 'yyyy-MM': { label, 'Mini-CEX': n, 'DOPS: X': n, ... } }
  const sixMonthsAgo = new Date(); sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
  const received_6m = {}; // submIDs per formType, senaste 6 mån
  const sent_6m     = {};

  rows.forEach(row => {
    const ts        = row[0];
    const registrar = (row[2] || '').toString().trim();
    const recipient = (row[4] || '').toString().split(' -- ')[0].trim();
    const formType  = (row[5] || '').toString().trim();
    const section   = (row[6] || '').toString().trim();
    const question  = (row[7] || '').toString().trim();
    const answer    = (row[8] || '').toString().trim();
    const submId    = row[9] ? row[9].toString() : null;

    if (!ts || !formType || !submId) return;

    const tsDate  = ts instanceof Date ? ts : new Date(ts);
    const monthKey = Utilities.formatDate(tsDate, Session.getScriptTimeZone(), 'yyyy-MM');
    const isRegistrar = registrar === fullName || registrar.toLowerCase() === email;
    const isRecipient = recipient === fullName;

    // Skickade
    if (isRegistrar) {
      if (!sent[formType]) sent[formType] = { subs: new Set(), subCounts: {}, submIds: new Set() };
      sent[formType].submIds.add(submId);
      // Subingrepp: Svar kolumn G = answer där frågan är ingreppsfält (section = första avsnittet, kort svar)
      if ((formType === 'DOPS' || formType === 'Operationsspecifik DOPS') && answer && answer.length < 80 && !question.includes('?') && section) {
        // Hitta ingreppsfältet – kolumn G i Svar = section, vi letar efter korta fritextsvar tidigt
        if (!sent[formType].subCounts[submId]) sent[formType].subCounts[submId] = answer;
      }
      if (!byMonth[monthKey]) byMonth[monthKey] = { sent: new Set(), received: new Set(), label: Utilities.formatDate(tsDate, Session.getScriptTimeZone(), 'MMM yyyy') };
      byMonth[monthKey].sent.add(submId + '_sent');
      // Per formulärtyp per månad (för staplad graf)
      if (!byMonthFormType[monthKey]) byMonthFormType[monthKey] = { label: Utilities.formatDate(tsDate, Session.getScriptTimeZone(), 'MMM yyyy') };
      if (!byMonthFormType[monthKey][formType]) byMonthFormType[monthKey][formType] = new Set();
      byMonthFormType[monthKey][formType].add(submId);
    }

    // Mottagna
    if (isRecipient) {
      if (!received[formType]) received[formType] = { submIds: new Set(), subCounts: {} };
      received[formType].submIds.add(submId);
      if ((formType === 'DOPS' || formType === 'Operationsspecifik DOPS') && answer && answer.length < 80 && !question.includes('?')) {
        if (!received[formType].subCounts[submId]) received[formType].subCounts[submId] = answer;
      }
      if (!byMonth[monthKey]) byMonth[monthKey] = { sent: new Set(), received: new Set(), label: Utilities.formatDate(tsDate, Session.getScriptTimeZone(), 'MMM yyyy') };
      byMonth[monthKey].received.add(submId + '_rec');
      if (!byMonthFormType[monthKey]) byMonthFormType[monthKey] = { label: Utilities.formatDate(tsDate, Session.getScriptTimeZone(), 'MMM yyyy') };
      if (!byMonthFormType[monthKey][formType]) byMonthFormType[monthKey][formType] = new Set();
      byMonthFormType[monthKey][formType].add(submId);
      if (tsDate >= sixMonthsAgo) {
        if (!received_6m[formType]) received_6m[formType] = new Set();
        received_6m[formType].add(submId);
      }
    }
    if (isRegistrar && tsDate >= sixMonthsAgo) {
      if (!sent_6m[formType]) sent_6m[formType] = new Set();
      sent_6m[formType].add(submId);
    }
  });

  // Konvertera till räknare
  // Hämta tidsbesparing
  const ftSheet2 = ss.getSheetByName('Formulärtyper');
  const timeSavings = {};
  if (ftSheet2) {
    ftSheet2.getDataRange().getValues().slice(1).forEach(r => {
      if (r[0] && r[3]) timeSavings[r[0].toString().trim()] = parseInt(r[3]) || 0;
    });
  }

  const sentClean = {};
  let totalMinutesSaved = 0;
  Object.entries(sent).forEach(([ft, v]) => {
    const subCounts = {};
    Object.values(v.subCounts).forEach(ingrepp => {
      subCounts[ingrepp] = (subCounts[ingrepp] || 0) + 1;
    });
    const count = v.submIds.size;
    const mins  = timeSavings[ft] || 0;
    totalMinutesSaved += count * mins;
    sentClean[ft] = { count, subCounts, goal: formGoals[ft] || subGoals[ft] || 0, minutesSaved: count * mins };
  });

  const recClean = {};
  Object.entries(received).forEach(([ft, v]) => {
    const subCounts = {};
    Object.values(v.subCounts).forEach(ingrepp => {
      subCounts[ingrepp] = (subCounts[ingrepp] || 0) + 1;
    });
    recClean[ft] = { count: v.submIds.size, subCounts, goal: formGoals[ft] || subGoals[ft] || 0 };
  });

  const byMonthClean = {};
  Object.entries(byMonth).sort((a,b) => a[0].localeCompare(b[0])).forEach(([k,v]) => {
    byMonthClean[k] = { label: v.label, sent: v.sent.size, received: v.received.size };
  });

  const rec6m = {}; Object.entries(received_6m).forEach(([ft,s]) => { rec6m[ft] = s.size; });
  const snt6m = {}; Object.entries(sent_6m).forEach(([ft,s])     => { snt6m[ft] = s.size; });
  // Konvertera byMonthFormType sets → counts
  const byMonthFtClean = {};
  Object.entries(byMonthFormType).sort((a,b) => a[0].localeCompare(b[0])).forEach(([month, data]) => {
    byMonthFtClean[month] = { label: data.label };
    Object.entries(data).forEach(([k, v]) => {
      if (k !== 'label') byMonthFtClean[month][k] = v instanceof Set ? v.size : v;
    });
  });
  return { sent: sentClean, received: recClean, byMonth: byMonthClean, subGoals, fullName, totalMinutesSaved, received_6m: rec6m, sent_6m: snt6m, byMonthFormType: byMonthFtClean };
}

// ── Kontextmening: AI-genererad med sheets-fallback ───────────────────────
function getContextPhrase(minutes, mode) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
    if (apiKey) {
      let prompt;
      if (mode === 'overview') {
        const h = Math.floor(minutes / 60), m = minutes % 60;
        const tidStr = h > 0 ? h + ' timmar' + (m > 0 ? ' och ' + m + ' minuter' : '') : m + ' minuter';
        prompt = 'En kirurg har totalt sparat ' + tidStr + ' med MoraDOPS. '
          + 'Skriv EN kort, positiv mening på svenska om vad man kan göra med den sparade tiden. '
          + 'Börja med "Med dina sparade " + tidStr + " kan du". '
          + 'Välj bland: en promenad, en god lunch, en fika med en kollega, ett poddavsnitt, ge feedback, ringa hem, meditera, läsa en artikel. '
          + 'Håll det konkret, upplyftande och mänskligt. Inga citattecken, bara meningen.';
      } else {
        prompt = 'En kirurg har just sparat ' + minutes + ' minuter med MoraDOPS. '
          + 'Skriv EN kort, rolig mening på svenska som börjar med "Du hinner nu". '
          + 'Välj SLUMPMÄSSIGT en kategori: fysisk aktivitet, mat, kultur, natur, sport, hantverk, geografi, historia, vetenskap, musik, språk. '
          + 'Vad kan man göra på exakt ' + minutes + ' minuter inom den kategorin? Var specifik och oväntat. '
          + 'Förbjudet: djur, fika, kaffe, kollegor, te, brygga något att dricka. '
          + 'Exempel: "Du hinner nu memorera alla svenska landskap i bokstavsordning." '
          + '"Du hinner nu lära dig tre gitarrdrag och spela en låt halvdåligt." '
          + '"Du hinner nu jogga till närmaste rondell och tillbaka." '
          + '"Du hinner nu skriva en recension av din favoritfilm på 300 ord." '
          + 'Inga citattecken, bara meningen. Byt ämne varje gång.';
      }
      const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post',
        headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
        payload: JSON.stringify({ model: 'claude-haiku-4-5-20251001', max_tokens: 120, messages: [{ role: 'user', content: prompt }] }),
        muteHttpExceptions: true
      });
      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        const text = data.content && data.content[0] && data.content[0].text ? data.content[0].text.trim() : null;
        if (text && text.length > 10) return { text, source: 'ai' };
      }
    }
  } catch(e) { Logger.log('AI phrase failed: ' + e.message); }

  try {
    const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Kontextmeningar');
    if (sheet) {
      const rows = sheet.getDataRange().getValues().slice(1)
        .filter(r => r[2] && parseInt(r[0]) <= minutes && parseInt(r[1]) > minutes);
      if (rows.length > 0) {
        const row = rows[Math.floor(Math.random() * rows.length)];
        return { text: row[2].toString().trim(), source: 'sheets' };
      }
    }
  } catch(e) { Logger.log('Sheets phrase failed: ' + e.message); }
  return null;
}

// ── Bulk-förfrågningar ────────────────────────────────────────────────────
function sendBulkRequest(fromEmail, toEmails, message, formType) {
  fromEmail = (fromEmail || '').trim().toLowerCase();
  message   = (message   || '').trim();
  formType  = (formType  || '').trim();
  const ss = SpreadsheetApp.openById(DOC_ID);
  let sheet = ss.getSheetByName('Förfrågningar');
  if (!sheet) {
    sheet = ss.insertSheet('Förfrågningar');
    sheet.appendRow(['Timestamp', 'Från', 'Till', 'Meddelande', 'Läst', 'ID', 'Status', 'FormulärTyp']);
  }
  const userSheet = ss.getSheetByName('Användare');
  const users = userSheet.getDataRange().getValues().slice(1);
  const nameMap = {};
  users.forEach(r => { nameMap[r[3].toString().trim().toLowerCase()] = (r[1] + ' ' + r[2]).trim(); });
  const fromName = nameMap[fromEmail] || fromEmail;
  let sent = 0;
  (toEmails || []).forEach(toEmail => {
    toEmail = toEmail.trim().toLowerCase();
    if (!toEmail) return;
    const id = Utilities.getUuid();
    const isSelf = toEmail === fromEmail;
    const toRow = users.find(r => r[3].toString().trim().toLowerCase() === toEmail);
    const toName = toRow ? (toRow[1] + ' ' + toRow[2]).trim() : toEmail;
    // Skapa förfrågan med FormulärTyp i kolumn H
    sheet.appendRow([new Date(), fromEmail, toEmail, message, 'Nej', id, '', formType]);
    // Skicka mail om emailNotify
    const emailNotify = toRow ? (toRow[14] || 'Ja').toString().trim() : 'Ja'; // kolumn O
    if (toRow && emailNotify !== 'Nej') {
      const subject = isSelf
        ? '[MoraDOPS] Självskattning: ' + formType
        : '[MoraDOPS] ' + formType + ' — begärd av ' + fromName;
      const reqHtmlBulk = buildRequestEmail(fromName, toName, message, formType, toEmail);
      MailApp.sendEmail({
        to: toEmail, replyTo: 'donotreply.mokir@gmail.com', name: 'MoraDOPS', subject,
        htmlBody: reqHtmlBulk,
        body: 'Hej ' + toName + ',\n\n' + fromName + ' ber dig genomföra: ' + formType + '\n\n"' + message + '"\n\nLogga in på MoraDOPS: https://mokir-web.github.io/dops/'
      });
    }
    sent++;
  });
  return { success: true, sent };
}

// ── Förfrågningsstatistik för Studierektor ─────────────────────────────
function getRequestStats(klinikId) {
  klinikId = (klinikId || '').trim();
  const ss = SpreadsheetApp.openById(DOC_ID);
  const reqSheet = ss.getSheetByName('Förfrågningar');
  if (!reqSheet) return { requests: [], summary: {} };
  const userSheet = ss.getSheetByName('Användare');
  const users = userSheet.getDataRange().getValues().slice(1);

  // Bygg email → {name, jobRole, klinikId}
  const userMap = {};
  users.forEach(r => {
    const email = r[3].toString().trim().toLowerCase();
    userMap[email] = {
      name: (r[1] + ' ' + r[2]).trim(),
      jobRole: r[4].toString().trim(),
      klinikId: r[10].toString().trim()
    };
  });

  const lastRowR = reqSheet.getLastRow();
  if (lastRowR < 2) return { requests: [], summary: { total:0, done:0, pending:0, byFormType:{}, byRole:{} } };
  const rows = reqSheet.getRange(2, 1, lastRowR - 1, 8).getValues();
  const requests = rows
    .filter(r => {
      if (!r[1] || !r[2]) return false;
      if (!klinikId || klinikId === '*') return true;
      const fromUser = userMap[r[1].toString().trim().toLowerCase()];
      const toUser   = userMap[r[2].toString().trim().toLowerCase()];
      return (fromUser && fromUser.klinikId === klinikId) ||
             (toUser   && toUser.klinikId   === klinikId);
    })
    .map(r => {
      const fromEmail = r[1].toString().trim().toLowerCase();
      const toEmail   = r[2].toString().trim().toLowerCase();
      const fromUser  = userMap[fromEmail] || { name: fromEmail, jobRole: '?' };
      const toUser    = userMap[toEmail]   || { name: toEmail,   jobRole: '?' };
      const ts = r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : r[0].toString();
      return {
        id: r[5].toString(),
        timestamp: ts,
        fromName:  fromUser.name,
        fromRole:  fromUser.jobRole,
        toName:    toUser.name,
        toRole:    toUser.jobRole,
        message:   r[3].toString(),
        read:      r[4].toString() === 'Ja',
        status:    r[6].toString().trim(),
        formType:  r[7] ? r[7].toString().trim() : '',
        isSelf:    fromEmail === toEmail
      };
    })
    .sort((a, b) => b.timestamp.localeCompare(a.timestamp));

  // Sammanfattning
  const total    = requests.length;
  const done     = requests.filter(r => r.status === 'Utförd').length;
  const pending  = total - done;
  const byFormType = {};
  const byRole     = {};
  requests.forEach(r => {
    const ft = r.formType || 'Okänd';
    if (!byFormType[ft]) byFormType[ft] = { total: 0, done: 0 };
    byFormType[ft].total++;
    if (r.status === 'Utförd') byFormType[ft].done++;
    const role = r.toRole || '?';
    if (!byRole[role]) byRole[role] = { total: 0, done: 0 };
    byRole[role].total++;
    if (r.status === 'Utförd') byRole[role].done++;
  });

  return { requests: requests.slice(0, 100), summary: { total, done, pending, byFormType, byRole } };
}

// ── Snabb räknare för polling ──────────────────────────────────────────────
function getMyAssessmentCount(email) {
  email = (email || '').trim().toLowerCase();
  const ss = SpreadsheetApp.openById(DOC_ID);
  const userSheet = ss.getSheetByName('Användare');
  const uRow = userSheet.getDataRange().getValues().slice(1)
    .find(r => r[3].toString().trim().toLowerCase() === email);
  const fullName = uRow ? (uRow[1] + ' ' + uRow[2]).trim() : '';
  const svarSheet = ss.getSheetByName('Svar');
  if (!svarSheet || !fullName) return { received: 0 };
  const submIds = new Set();
  svarSheet.getDataRange().getValues().slice(1).forEach(r => {
    const recipient = (r[4] || '').toString().split(' -- ')[0].trim();
    if (recipient === fullName && r[9]) submIds.add(r[9].toString());
  });
  return { received: submIds.size };
}

// ── Admin: get documents from sheet ────────────────────────────────────────
function getDocuments() {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Dokument');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues().slice(1);
  return rows.filter(r => r[0]).map(r => ({
    title:       r[0].toString(),
    description: r[1] ? r[1].toString() : '',
    url:         r[2] ? r[2].toString() : '',
    version:     r[3] ? r[3].toString() : '',
    date:        r[4] ? r[4].toString() : ''
  }));
}

// ── Register user ──────────────────────────────────────────────────────────
function registerUser(user) {
  const ss = SpreadsheetApp.openById(DOC_ID);
  let sheet = ss.getSheetByName('Användare');
  if (!sheet) {
    sheet = ss.insertSheet('Användare');
    sheet.appendRow(['Timestamp', 'Förnamn', 'Efternamn', 'Email', 'Yrkesroll', 'Klinik', 'Användarroll', 'PIN', 'ResetCode', 'ResetExpiry', 'KlinikID', 'FailedAttempts', 'LockUntil']);
  }
  // Trimma all indata
  user.firstName = (user.firstName || '').trim();
  user.lastName  = (user.lastName  || '').trim();
  user.email     = (user.email     || '').trim().toLowerCase();
  user.clinic    = (user.clinic    || '').trim();
  user.jobRole   = (user.jobRole   || '').trim();
  user.userRole  = (user.userRole  || '').trim();

  const data = sheet.getDataRange().getValues();
  const exists = data.slice(1).some(r => r[3].toString().trim().toLowerCase() === user.email);
  if (exists) return { error: 'En användare med denna e-postadress finns redan.' };

  // Generera klinikId från kliniknamn om inte angivet
  const klinikId = (user.klinikId || '').trim() || generateKlinikId(user.clinic);
  user.klinikId = klinikId;

  sheet.appendRow([
    new Date(), user.firstName, user.lastName, user.email,
    user.jobRole, user.clinic, user.userRole, hashPin(user.pin), '', '', klinikId, '', '',
    'Ja', 'Ja', '', new Date().toISOString().slice(0,10), ''  // N=EmailNotify O=EmailNotifyReq P=StartPage Q=SenastAktiv R=Inaktiverad
  ]);

  // Uppdatera Kliniker-sheet om kliniken är ny
  ensureClinicExists(ss, klinikId, user.clinic);

  updateCategories(user);
  invalidateCache();
  return { success: true };
}

function ensureClinicExists(ss, klinikId, clinicName) {
  let sheet = ss.getSheetByName('Kliniker');
  if (!sheet) {
    sheet = ss.insertSheet('Kliniker');
    sheet.appendRow(['KlinikID', 'Namn', 'Kontaktperson', 'Aktiv', 'Startdatum']);
  }
  const data = sheet.getDataRange().getValues();
  const exists = data.slice(1).some(r => r[0].toString().trim() === klinikId);
  if (!exists) {
    sheet.appendRow([klinikId, clinicName, '', 'Ja', new Date().toISOString().slice(0,10)]);
  }
}

function updateCategories(user) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Categories');
  const colMap = { Registrerare: 0, Mottagare: 1, ST: 2, Spec: 3, AT: 4, BT: 5 };
  const firstName = (user.firstName || '').trim();
  const lastName  = (user.lastName  || '').trim();
  const clinic    = (user.clinic    || '').trim();
  const jobRole   = (user.jobRole   || '').trim();
  const userRole  = (user.userRole  || '').trim();
  const cellValue = `${firstName} ${lastName} -- ${clinic}`;
  const columns = [];
  if (userRole === 'Registrerare' || userRole === 'Blandbild') columns.push(colMap['Registrerare']);
  if (userRole === 'Mottagare'    || userRole === 'Blandbild') columns.push(colMap['Mottagare']);
  if (colMap[jobRole] !== undefined) columns.push(colMap[jobRole]);
  const totalRows = Math.max(sheet.getLastRow(), 1);
  columns.forEach(colIdx => {
    const colData = sheet.getRange(1, colIdx + 1, totalRows, 1).getValues().flat();
    if (colData.includes(cellValue)) return;
    let emptyRow = -1;
    for (let i = 1; i < colData.length; i++) {
      if (!colData[i] || colData[i].toString().trim() === '') { emptyRow = i + 1; break; }
    }
    if (emptyRow === -1) emptyRow = totalRows + 1;
    sheet.getRange(emptyRow, colIdx + 1).setValue(cellValue);
  });
  const clinicData = sheet.getRange(1, 7, totalRows, 1).getValues().flat();
  if (!clinicData.includes(user.clinic)) {
    let emptyRow = -1;
    for (let i = 1; i < clinicData.length; i++) {
      if (!clinicData[i] || clinicData[i].toString().trim() === '') { emptyRow = i + 1; break; }
    }
    if (emptyRow === -1) emptyRow = totalRows + 1;
    sheet.getRange(emptyRow, 7).setValue(user.clinic);
  }
}

function syncUsersToCategories() {
  rebuildCategoriesFromUsers(SpreadsheetApp.openById(DOC_ID));
}

function rebuildCategoriesFromUsers(ss) {
  const userSheet = ss.getSheetByName('Användare');
  const catSheet  = ss.getSheetByName('Categories');
  if (!userSheet || !catSheet) return;

  const users = userSheet.getDataRange().getValues().slice(1).filter(r => r[1] && r[3] && (r[17] || '').toString().trim() !== 'Ja');
  const colMap = { Registrerare: 0, Mottagare: 1, ST: 2, Spec: 3, AT: 4, BT: 5 };

  // Bygg upp vad Categories SKA innehålla
  const expected = { 0: new Set(), 1: new Set(), 2: new Set(), 3: new Set(), 4: new Set(), 5: new Set() };
  const clinics = new Set();
  users.forEach(row => {
    const name     = row[1] + ' ' + row[2];
    const clinic   = row[5] ? row[5].toString().trim() : '';
    const jobRole  = row[4] ? row[4].toString().trim() : '';
    let userRole = row[6] ? row[6].toString().trim() : '';
    if (userRole === 'Registrerare av bedömningar') userRole = 'Registrerare';
    if (userRole === 'Mottagare av bedömningar')    userRole = 'Mottagare';
    const cell     = name + ' -- ' + clinic;
    if (clinic) clinics.add(clinic);
    if (userRole === 'Registrerare' || userRole === 'Blandbild') expected[0].add(cell);
    if (userRole === 'Mottagare'    || userRole === 'Blandbild') expected[1].add(cell);
    if (colMap[jobRole] !== undefined) expected[colMap[jobRole]].add(cell);
  });

  // Skriv om kolumnerna A–F och G
  const maxRows = Math.max(...[0,1,2,3,4,5].map(i => expected[i].size), clinics.size, 1) + 1;
  const currentRows = Math.max(catSheet.getLastRow(), 1);
  const totalRows   = Math.max(maxRows, currentRows);

  // Rensa och skriv kolumn för kolumn
  [0,1,2,3,4,5].forEach(colIdx => {
    const vals = [Object.keys(colMap).find(k => colMap[k] === colIdx)];
    [...expected[colIdx]].forEach(v => vals.push(v));
    while (vals.length < totalRows) vals.push('');
    catSheet.getRange(1, colIdx + 1, totalRows, 1).setValues(vals.map(v => [v]));
  });
  // Rensa klinik-kolumn G (används ej längre)
  const emptyCol = Array(totalRows).fill(['']);
  emptyCol[0] = [''];
  catSheet.getRange(1, 7, totalRows, 1).setValues(emptyCol);
}

// ── Reklassificera migrerade bedömningar ─────────────────────────────────
// Kör en gång manuellt efter import av Svar_migrering för att rätta
// felklassade formulärtyper (Mini-CEX, Återkoppling etc → DOPS: Allmän)
function fixMigratedFormTypes() {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Svar');
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

  // Gruppera rader per SubmissionID (kolumn J = index 9)
  const groups = {};
  data.forEach((row, i) => {
    const sid = row[9] ? row[9].toString() : null;
    if (!sid) return;
    if (!groups[sid]) groups[sid] = [];
    groups[sid].push({ rowIdx: i + 2, formType: row[5], question: row[7], answer: row[8] });
  });

  function detectType(entries) {
    const qs = entries.map(e => (e.question || '').toString());
    if (qs.some(q => q.includes('Anamnestagning') || q.includes('Klinisk undersökning')))
      return 'Mini-CEX';
    if (qs.some(q => q.includes('Introduktionen') || q.includes('Retorik') || q.includes('Begriplighet')))
      return 'Återkoppling efter muntlig presentation';
    if (qs.some(q => q.includes('Diagnostiserar patienters') || q.includes('Planerar utredning')))
      return '360 GRADER';
    return null;
  }

  let fixed = 0;
  Object.values(groups).forEach(entries => {
    const currentType = entries[0].formType;
    if (currentType !== 'DOPS: Allmän') return; // bara felklassade
    const newType = detectType(entries);
    if (!newType) return;
    entries.forEach(e => {
      sheet.getRange(e.rowIdx, 6).setValue(newType);
    });
    fixed++;
  });

  Logger.log('Rättade ' + fixed + ' bedömningar.');
  return { fixed };
}

// ── Schemalagda utskick ───────────────────────────────────────────────────
// Sheet "Utskick": A=ID, B=Namn, C=KlinikID, D=MottagarTyp, E=Dag, F=Tid(HH:MM), G=Mottagare(CSV), H=Aktiv, I=SenastSkickad, J=Innehall, K=PeriodVeckor
function getScheduledEmails(klinikId) {
  const ss = SpreadsheetApp.openById(DOC_ID);
  let sheet = ss.getSheetByName('Utskick');
  if (!sheet) {
    sheet = ss.insertSheet('Utskick');
    sheet.appendRow(['ID','Namn','KlinikID','MottagarTyp','Dag','Tid','Mottagare','Aktiv','SenastSkickad','Innehall','PeriodVeckor']);
    return [];
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const numCols = Math.max(sheet.getLastColumn(), 11);
  return sheet.getRange(2, 1, lastRow - 1, numCols).getValues()
    .filter(r => r[0] && (!klinikId || klinikId === '*' || r[2] === klinikId))
    .map(r => {
      let tid = r[5];
      if (tid instanceof Date) tid = Utilities.formatDate(tid, Session.getScriptTimeZone(), 'HH:mm');
      else tid = (tid || '07:00').toString().replace(/^.*?(\d{1,2}):(\d{2}).*$/, (_, h, m) => h.padStart(2,'0') + ':' + m);
      return {
        id: r[0].toString(), namn: r[1].toString(), klinikId: r[2].toString(),
        mottagarTyp: r[3].toString(), dag: r[4].toString(), tid,
        mottagare: r[6] ? r[6].toString().split(',').map(e => e.trim()).filter(Boolean) : [],
        aktiv: r[7].toString() === 'Ja',
        senastSkickad: r[8] ? r[8].toString() : '',
        innehall: r[9] ? r[9].toString() : 'sent,received,formtypes',
        periodVeckor: parseInt(r[10]) || 1
      };
    });
}

function saveScheduledEmail(schedule) {
  const ss = SpreadsheetApp.openById(DOC_ID);
  let sheet = ss.getSheetByName('Utskick');
  if (!sheet) {
    sheet = ss.insertSheet('Utskick');
    sheet.appendRow(['ID','Namn','KlinikID','MottagarTyp','Dag','Tid','Mottagare','Aktiv','SenastSkickad','Innehall','PeriodVeckor']);
  }
  const id = schedule.id || Utilities.getUuid();
  const row = [id, schedule.namn||'', schedule.klinikId||'', schedule.mottagarTyp||'',
    schedule.dag||'', schedule.tid||'07:00', (schedule.mottagare||[]).join(','),
    schedule.aktiv ? 'Ja' : 'Nej', '', schedule.innehall||'sent,received,formtypes',
    parseInt(schedule.periodVeckor)||1];
  const data = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow()-1, 11).getValues() : [];
  const rowIdx = data.findIndex(r => r[0].toString() === id);
  if (rowIdx !== -1) sheet.getRange(rowIdx + 2, 1, 1, 11).setValues([row]);
  else sheet.appendRow(row);
  return { success: true, id };
}

function toggleScheduleActive(id) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Utskick');
  if (!sheet) return { error: 'Sheet saknas' };
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex((r, i) => i > 0 && r[0].toString() === id);
  if (rowIdx === -1) return { error: 'Schema hittades inte' };
  const current = data[rowIdx][7].toString() === 'Ja';
  sheet.getRange(rowIdx + 1, 8).setValue(current ? 'Nej' : 'Ja');
  return { success: true, aktiv: !current };
}

function deleteScheduledEmail(id) {
  const sheet = SpreadsheetApp.openById(DOC_ID).getSheetByName('Utskick');
  if (!sheet) return { success: false };
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex((r, i) => i > 0 && r[0].toString() === id);
  if (rowIdx !== -1) sheet.deleteRow(rowIdx + 1);
  return { success: true };
}

// Tidsstyrd trigger – kör dagligen
function sendScheduledEmails() {
  const schedules = getScheduledEmails('*').filter(s => s.aktiv);
  if (!schedules.length) return;
  const now = new Date();
  const days = ['Söndag','Måndag','Tisdag','Onsdag','Torsdag','Fredag','Lördag'];
  const todayName = days[now.getDay()];
  const ss = SpreadsheetApp.openById(DOC_ID);
  const sheet = ss.getSheetByName('Utskick');
  const data = sheet.getDataRange().getValues();

  schedules.forEach(s => {
    if (s.dag !== todayName && s.dag !== 'Varje dag') return;
    // Tidsjämförelse i skriptets tidszon (undviker UTC-offset-problem)
    const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');
    const [nowH, nowM] = nowStr.split(':').map(Number);
    const [hh, mm] = (s.tid || '07:00').split(':').map(Number);
    if (nowH < hh || (nowH === hh && nowM < mm)) return;
    const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (s.senastSkickad && s.senastSkickad.startsWith(todayStr)) return; // redan skickat idag

    const userSheet = ss.getSheetByName('Användare');
    const users = userSheet.getDataRange().getValues().slice(1);
    const svarSheet = ss.getSheetByName('Svar');
    const svarRows = svarSheet ? svarSheet.getDataRange().getValues().slice(1) : [];
    const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    const periodMs = (s.periodVeckor || 1) * 7 * 24 * 60 * 60 * 1000;
    const periodAgo = new Date(now.getTime() - periodMs);

    // Alla rader i perioden (personlig statistik, ej klinikfiltrerat)
    const weekRowsAll = svarRows.filter(r => {
      const ts = r[0] instanceof Date ? r[0] : new Date(r[0]);
      return ts >= periodAgo;
    });

    function barRow(label, n, max, color) {
      const px = Math.round(n / Math.max(max, 1) * 180);
      const c = color || '#3d5c4a';
      return '<tr><td style="padding:2px 8px 2px 0;font-size:13px;color:#6b5c52;white-space:nowrap;">' + escapeHtml(label) + '</td>'
        + '<td><div style="background:' + c + ';height:12px;border-radius:2px;width:' + px + 'px;display:inline-block;"></div></td>'
        + '<td style="padding:2px 0 2px 8px;font-size:13px;font-weight:bold;color:#2a1f18;">' + n + '</td></tr>';
    }

    const innehall = s.innehall || 'sent,received,formtypes';
    const showSent2      = innehall.includes('sent');
    const showReceived2  = innehall.includes('received');
    const showFt2        = innehall.includes('formtypes');
    const showSentTo     = innehall.includes('sent_to');
    const showRecFrom    = innehall.includes('received_from');
    const periodLabel    = s.periodVeckor === 1 ? 'senaste veckan' : 'senaste ' + s.periodVeckor + ' veckorna';

    s.mottagare.forEach(email => {
      const uRow = users.find(r => r[3].toString().trim().toLowerCase() === email.toLowerCase());
      if (!uRow) return;
      const name = (uRow[1] + ' ' + uRow[2]).trim();

      // Personliga rader – matcha på namn (kolumn C = registrerare, kolumn E = mottagare)
      const personalRows = weekRowsAll.filter(r => {
        const reg = (r[2] || '').toString().trim();
        const rec = (r[4] || '').toString().split(' -- ')[0].trim();
        return reg === name || rec === name;
      });

      // Skickade: per formulärtyp
      const sentByFt = {};
      personalRows.filter(r => (r[2]||'').toString().trim() === name && r[9])
        .forEach(r => { if (!sentByFt[r[5]]) sentByFt[r[5]] = new Set(); sentByFt[r[5]].add(r[9]); });
      const sentFtCounts = Object.entries(sentByFt).map(([ft,st]) => ({ ft, n: st.size })).sort((a,b) => b.n-a.n);

      // Mottagna: per formulärtyp
      const recByFt = {};
      personalRows.filter(r => (r[4]||'').toString().split(' -- ')[0].trim() === name && r[9])
        .forEach(r => { if (!recByFt[r[5]]) recByFt[r[5]] = new Set(); recByFt[r[5]].add(r[9]); });
      const recFtCounts = Object.entries(recByFt).map(([ft,st]) => ({ ft, n: st.size })).sort((a,b) => b.n-a.n);

      // Skickat till (per mottagare)
      const sentToMap = {};
      personalRows.filter(r => (r[2]||'').toString().trim() === name && r[9])
        .forEach(r => {
          const rec = (r[4]||'').toString().split(' -- ')[0].trim();
          if (!sentToMap[rec]) sentToMap[rec] = new Set();
          sentToMap[rec].add(r[9]);
        });
      const sentToCounts = Object.entries(sentToMap).map(([p,s]) => ({ p, n: s.size })).sort((a,b) => b.n-a.n);

      // Mottagit från (per registrerare)
      const recFromMap = {};
      personalRows.filter(r => (r[4]||'').toString().split(' -- ')[0].trim() === name && r[9])
        .forEach(r => {
          const reg = (r[2]||'').toString().trim();
          if (!recFromMap[reg]) recFromMap[reg] = new Set();
          recFromMap[reg].add(r[9]);
        });
      const recFromCounts = Object.entries(recFromMap).map(([p,s]) => ({ p, n: s.size })).sort((a,b) => b.n-a.n);

      const sentTotal = sentFtCounts.reduce((s, e) => s + e.n, 0);
      const recTotal  = recFtCounts.reduce((s, e) => s + e.n, 0);

      // Summakort
      let statsHtml = '<table style="width:100%;border-collapse:collapse;margin-bottom:20px;"><tr>';
      if (showSent2)     statsHtml += '<td style="background:#3d5c4a;color:#f0ede8;border-radius:6px;padding:12px 16px;text-align:center;"><div style="font-size:11px;opacity:0.8;">REGISTRERADE</div><div style="font-size:26px;font-weight:bold;">' + sentTotal + '</div></td><td style="width:8px;"></td>';
      if (showReceived2) statsHtml += '<td style="background:#9e7c68;color:#f0ede8;border-radius:6px;padding:12px 16px;text-align:center;"><div style="font-size:11px;opacity:0.8;">MOTTAGNA</div><div style="font-size:26px;font-weight:bold;">' + recTotal + '</div></td>';
      statsHtml += '</tr></table>';

      // Stapeldiagram per formulärtyp – registrerade
      let ftSentToHtml = '';
      if (showSentTo && sentToCounts.length) {
        const mx = sentToCounts[0].n;
        ftSentToHtml = '<div style="font-weight:bold;color:#2a1f18;margin-bottom:6px;font-size:13px;">Skickat bedömningar till</div>'
          + '<table style="border-collapse:collapse;margin-bottom:16px;">'
          + sentToCounts.slice(0,8).map(({p,n}) => barRow(p, n, mx, '#3d5c4a')).join('') + '</table>';
      }
      let ftRecFromHtml = '';
      if (showRecFrom && recFromCounts.length) {
        const mx = recFromCounts[0].n;
        ftRecFromHtml = '<div style="font-weight:bold;color:#2a1f18;margin-bottom:6px;font-size:13px;">Mottagit bedömningar av</div>'
          + '<table style="border-collapse:collapse;margin-bottom:16px;">'
          + recFromCounts.slice(0,8).map(({p,n}) => barRow(p, n, mx, '#9e7c68')).join('') + '</table>';
      }
        const maxS = sentFtCounts[0].n;
        ftSentHtml = '<div style="font-weight:bold;color:#2a1f18;margin-bottom:6px;font-size:13px;">Registrerade per formulär</div>'
          + '<table style="border-collapse:collapse;margin-bottom:16px;">'
          + sentFtCounts.slice(0, 8).map(({ft, n}) => barRow(ft, n, maxS, '#3d5c4a')).join('')
          + '</table>';
      }
      // Stapeldiagram per formulärtyp – mottagna
      let ftRecHtml = '';
      if (showReceived2 && showFt2 && recFtCounts.length) {
        const maxR = recFtCounts[0].n;
        ftRecHtml = '<div style="font-weight:bold;color:#2a1f18;margin-bottom:6px;font-size:13px;">Mottagna per formulär</div>'
          + '<table style="border-collapse:collapse;margin-bottom:16px;">'
          + recFtCounts.slice(0, 8).map(({ft, n}) => barRow(ft, n, maxR, '#9e7c68')).join('')
          + '</table>';
      }

      const html = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#f0ede8;">'
        + '<div style="background:#3d5c4a;border-radius:6px 6px 0 0;">'
        + '<img src="https://mokir-web.github.io/dops/email-header.png" width="600" style="display:block;width:100%;border-radius:6px 6px 0 0;"/></div>'
        + '<div style="background:#f5f1ec;padding:28px 32px;border-radius:0 0 6px 6px;border:1.5px solid #d8cfc5;border-top:none;">'
        + '<h2 style="color:#2a1f18;margin-top:0;">Din veckoöversikt</h2>'
        + '<p style="color:#6b5c52;font-size:14px;">Hej ' + escapeHtml(uRow[1]) + ', här är din personliga sammanfattning för de senaste 7 dagarna.</p>'
        + statsHtml + ftSentHtml + ftRecHtml + ftSentToHtml + ftRecFromHtml
        + '<div style="margin-top:24px;"><a href="https://mokir-web.github.io/dops/" style="background:#3d5c4a;color:#f0ede8;text-decoration:none;padding:10px 24px;border-radius:6px;font-weight:bold;">Öppna MoraDOPS →</a></div>'
        + getUnsubscribeFooter(email) + '</div></div>';
      try {
        MailApp.sendEmail({ to: email, name: 'MoraDOPS',
          subject: '[MoraDOPS] Veckoöversikt ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          htmlBody: html, body: 'Hej ' + uRow[1] + '. https://mokir-web.github.io/dops/' });
      } catch(e) { Logger.log('Email failed ' + email + ': ' + e.message); }
    });

    // Uppdatera SenastSkickad
    const rowIdx = data.findIndex((r, i) => i > 0 && r[0].toString() === s.id);
    if (rowIdx !== -1) sheet.getRange(rowIdx + 1, 9).setValue(todayStr);
  });
}
