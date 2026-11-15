/**
 * Sri Ayyappa Swamy Padi Pooja – RSVP backend (Google Apps Script)
 * -----------------------------------------------------------------
 * - Serves the invitation page (Index.html) as a Web App
 * - Saves RSVPs to the "RSVPs" sheet of the spreadsheet this script is bound to
 * - One row per mobile number: re-submitting with the same number UPDATES the row
 * - Password-protected admin data for the host dashboard
 * - Also accepts JSON POSTs, so the page can be hosted elsewhere (GitHub Pages, Netlify)
 *
 * Setup: see README.md. In short: run setup() once, set ADMIN_PASSWORD in
 * Project Settings → Script Properties, then Deploy → New deployment → Web app.
 */

const SHEET_NAME = 'RSVPs';
const HEADERS = ['RSVP ID', 'Submitted', 'Updated', 'Name', 'Mobile', 'Email',
                 'Attendance', 'Adults', 'Children', 'Message', 'Invited As', 'Update Count'];
const COL = HEADERS.reduce((m, h, i) => (m[h] = i, m), {});

/* ---------------- Web app entry points ---------------- */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const guest = String(p.guest || p.to || '').replace(/[<>"'`\\]/g, '').replace(/\s+/g, ' ').trim().slice(0, 60);
  const t = HtmlService.createTemplateFromFile('Index');
  // "<" is escaped so nothing in the data can close the <script> tag
  t.boot = JSON.stringify({ appUrl: ScriptApp.getService().getUrl(), guest: guest }).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle('Sri Ayyappa Swamy Padi Pooja – Invitation')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** JSON API for pages hosted outside Apps Script. Body: {action:'rsvp', data:{...}} or {action:'admin', password:'...'} */
function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'rsvp') out = Object.assign({ ok: true }, submitRsvp(body.data || {}));
    else if (body.action === 'admin') out = Object.assign({ ok: true }, getAdminData(body.password));
    else out = { ok: false, error: 'Unknown action' };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- RSVP ---------------- */

function submitRsvp(d) {
  d = d || {};
  if (d.website) return { status: 'created', id: 'AYP-OK' };   // honeypot: quietly ignore bots

  const name = clean_(d.name, 80);
  const mobile = clean_(d.mobile, 20);
  const digits = mobile.replace(/\D/g, '');
  const email = clean_(d.email, 100);
  const attendance = ['Yes', 'No', 'Maybe'].indexOf(d.attendance) >= 0 ? d.attendance : '';
  let adults = toInt_(d.adults), children = toInt_(d.children);
  const message = clean_(d.message, 500);
  const invitedAs = clean_(d.invitedAs, 60);

  if (name.length < 2) throw new Error('Please enter your full name');
  if (digits.length < 10 || digits.length > 15) throw new Error('Please enter a valid mobile number');
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) throw new Error('Please check your email address');
  if (!attendance) throw new Error('Please choose Yes, Maybe or No');
  if (adults < 0 || adults > 30 || children < 0 || children > 30) throw new Error('Guest counts must be between 0 and 30');
  if (attendance === 'No') { adults = 0; children = 0; }
  else if (adults + children < 1) throw new Error('Please include at least one guest');

  const key = digits.slice(-10);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getSheet_();
    const now = new Date();
    const last = sh.getLastRow();
    let rowIndex = -1, existing = null;
    if (last > 1) {
      const values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
      for (let i = 0; i < values.length; i++) {
        if (String(values[i][COL['Mobile']]).replace(/\D/g, '').slice(-10) === key) {
          rowIndex = i + 2; existing = values[i]; break;
        }
      }
    }
    if (existing) {
      const row = existing.slice();
      row[COL['Updated']] = now;
      row[COL['Name']] = safe_(name);
      row[COL['Mobile']] = safe_(mobile);
      row[COL['Email']] = safe_(email);
      row[COL['Attendance']] = attendance;
      row[COL['Adults']] = adults;
      row[COL['Children']] = children;
      row[COL['Message']] = safe_(message);
      if (invitedAs) row[COL['Invited As']] = safe_(invitedAs);
      row[COL['Update Count']] = (Number(existing[COL['Update Count']]) || 0) + 1;
      sh.getRange(rowIndex, 1, 1, HEADERS.length).setValues([row]);
      notifyHost_('updated', name, attendance, adults, children, message);
      return { status: 'updated', id: String(existing[COL['RSVP ID']]) };
    }
    const id = 'AYP-' + Utilities.getUuid().slice(0, 6).toUpperCase();
    const row = [];
    row[COL['RSVP ID']] = id;
    row[COL['Submitted']] = now;
    row[COL['Updated']] = now;
    row[COL['Name']] = safe_(name);
    row[COL['Mobile']] = safe_(mobile);
    row[COL['Email']] = safe_(email);
    row[COL['Attendance']] = attendance;
    row[COL['Adults']] = adults;
    row[COL['Children']] = children;
    row[COL['Message']] = safe_(message);
    row[COL['Invited As']] = safe_(invitedAs);
    row[COL['Update Count']] = 0;
    sh.appendRow(row);
    notifyHost_('new', name, attendance, adults, children, message);
    return { status: 'created', id: id };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- Host dashboard ---------------- */

function getAdminData(password) {
  const stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!stored) throw new Error('Admin password is not set yet (Script Properties → ADMIN_PASSWORD)');
  if (String(password || '') !== stored) { Utilities.sleep(1000); throw new Error('Incorrect password'); }

  const sh = getSheet_();
  const last = sh.getLastRow();
  const values = last > 1 ? sh.getRange(2, 1, last - 1, HEADERS.length).getValues() : [];
  const iso = v => (v instanceof Date) ? v.toISOString() : String(v || '');
  const rows = values.filter(r => r[COL['Name']] !== '').map(r => ({
    id: String(r[COL['RSVP ID']]),
    createdAt: iso(r[COL['Submitted']]),
    updatedAt: iso(r[COL['Updated']]),
    name: String(r[COL['Name']]),
    mobile: String(r[COL['Mobile']]),
    email: String(r[COL['Email']]),
    attendance: String(r[COL['Attendance']]),
    adults: Number(r[COL['Adults']]) || 0,
    children: Number(r[COL['Children']]) || 0,
    message: String(r[COL['Message']]),
    invitedAs: String(r[COL['Invited As']]),
    updates: Number(r[COL['Update Count']]) || 0
  }));

  const t = { responses: rows.length, yesFamilies: 0, yesAdults: 0, yesChildren: 0, maybeFamilies: 0, maybeGuests: 0, noFamilies: 0 };
  rows.forEach(r => {
    if (r.attendance === 'Yes') { t.yesFamilies++; t.yesAdults += r.adults; t.yesChildren += r.children; }
    else if (r.attendance === 'Maybe') { t.maybeFamilies++; t.maybeGuests += r.adults + r.children; }
    else if (r.attendance === 'No') { t.noFamilies++; }
  });
  t.confirmedGuests = t.yesAdults + t.yesChildren;
  return { rows: rows, totals: t };
}

/* ---------------- One-time setup ---------------- */

/** Run once from the Apps Script editor (select "setup" → Run) and accept the permissions. */
function setup() {
  const sh = getSheet_();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_PASSWORD')) {
    props.setProperty('ADMIN_PASSWORD', 'ChangeMe-Swamy18');
    Logger.log('Temporary admin password set to "ChangeMe-Swamy18". Change it in Project Settings → Script Properties.');
  }
  Logger.log('Sheet ready: ' + sh.getParent().getUrl());
}

/* ---------------- Helpers ---------------- */

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this script from Extensions → Apps Script inside your Google Sheet');
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#4a0a10').setFontColor('#f6d982');
    sh.setFrozenRows(1);
    sh.getRange('E:E').setNumberFormat('@');           // keep mobile numbers as text
    sh.getRange('B:C').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.setColumnWidth(COL['Message'] + 1, 320);
  }
  return sh;
}

function clean_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max); }
function toInt_(v) { const n = parseInt(v, 10); return isNaN(n) ? 0 : n; }
/** Stops spreadsheet formula injection (values starting with = + - @). */
function safe_(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; }

/** Optional: email the host on every RSVP. Add Script Property HOST_EMAIL to turn it on. */
function notifyHost_(kind, name, attendance, adults, children, message) {
  const to = PropertiesService.getScriptProperties().getProperty('HOST_EMAIL');
  if (!to) return;
  try {
    MailApp.sendEmail(to,
      'Padi Pooja RSVP ' + (kind === 'updated' ? '(updated)' : '') + ': ' + name + ' – ' + attendance,
      'Name: ' + name + '\nAttendance: ' + attendance + '\nAdults: ' + adults + '\nChildren: ' + children +
      (message ? '\nMessage: ' + message : '') + '\n\nSwamiye Saranam Ayyappa!');
  } catch (e) { console.warn('Email failed: ' + e); }
}
