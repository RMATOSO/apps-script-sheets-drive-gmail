/**
 * Sheets -> Drive -> Gmail automation (demo)
 *
 * When a row in the "Requests" tab is set to Status = "Approved":
 *   1. a Drive folder "<ID> - <Client>" is created inside the parent folder,
 *   2. a request summary file is saved in that folder,
 *   3. an e-mail notification is sent (test mode: only to the owner),
 *   4. the folder link and the notification time are written back to the row,
 *   5. every run is recorded in the "Log" tab.
 * A daily summary e-mail lists how many requests are in each status.
 *
 * Bound script: open the spreadsheet > Extensions > Apps Script.
 */

const REQ_SHEET = 'Requests';
const LOG_SHEET = 'Log';
const CFG_SHEET = 'Settings';
const TRIGGER_STATUS = 'Approved';

// 1-based column positions in the Requests tab
const COL = { ID: 1, CLIENT: 2, EMAIL: 3, REQUEST: 4, STATUS: 5, FOLDER: 6, NOTIFIED: 7 };

/* ---------- Menu ---------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Automation')
    .addItem('1. Set up (folder + triggers)', 'setup')
    .addItem('Process approved rows now', 'processAll')
    .addItem('Send daily summary now', 'sendDailySummary')
    .addSeparator()
    .addItem('Remove triggers', 'removeTriggers')
    .addToUi();
}

/* ---------- Setup ---------- */

function setup() {
  const settings = getSettings_();

  if (!settings.OWNER_EMAIL) {
    setSetting_('OWNER_EMAIL', Session.getEffectiveUser().getEmail());
  }
  if (!settings.PARENT_FOLDER_ID) {
    const folder = DriveApp.createFolder('Client Requests (demo)');
    setSetting_('PARENT_FOLDER_ID', folder.getId());
  }
  if (!settings.NOTIFY_MODE) setSetting_('NOTIFY_MODE', 'test');
  if (!settings.SUMMARY_HOUR) setSetting_('SUMMARY_HOUR', '8');

  removeTriggers();
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('onEditInstalled').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('sendDailySummary')
    .timeBased()
    .everyDays(1)
    .atHour(Number(getSettings_().SUMMARY_HOUR) || 8)
    .create();

  log_('-', 'setup', 'OK', 'Folder and triggers ready');
  SpreadsheetApp.getUi().alert(
    'Setup complete. Mode: ' + getSettings_().NOTIFY_MODE +
    '. In test mode every e-mail goes only to ' + getSettings_().OWNER_EMAIL + '.'
  );
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    ScriptApp.deleteTrigger(t);
  });
}

/* ---------- Triggers ---------- */

// Installable onEdit trigger (simple triggers cannot use Drive or Gmail)
function onEditInstalled(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== REQ_SHEET) return;
  if (e.range.getColumn() !== COL.STATUS || e.range.getRow() < 2) return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
  if (String(e.value).trim() !== TRIGGER_STATUS) return;
  processRow_(sh, e.range.getRow());
}

function processAll() {
  const sh = SpreadsheetApp.getActive().getSheetByName(REQ_SHEET);
  const last = sh.getLastRow();
  if (last < 2) return;
  const rows = sh.getRange(2, 1, last - 1, COL.NOTIFIED).getValues();
  rows.forEach(function (r, i) {
    if (String(r[COL.STATUS - 1]).trim() === TRIGGER_STATUS && !r[COL.NOTIFIED - 1]) {
      processRow_(sh, i + 2);
    }
  });
}

/* ---------- Core ---------- */

function processRow_(sh, row) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) {
    log_('-', 'process row ' + row, 'SKIPPED', 'Another run in progress');
    return;
  }
  let id = '-';
  try {
    const r = sh.getRange(row, 1, 1, COL.NOTIFIED).getValues()[0];
    id = String(r[COL.ID - 1]).trim();
    const client = String(r[COL.CLIENT - 1]).trim();
    const email = String(r[COL.EMAIL - 1]).trim();
    const request = String(r[COL.REQUEST - 1]).trim();

    // Idempotent: a fully processed row is skipped; a row whose folder exists
    // but whose e-mail failed only gets the notification again (no second folder)
    let folderUrl = String(r[COL.FOLDER - 1] || '').trim();
    if (folderUrl && r[COL.NOTIFIED - 1]) {
      log_(id, 'process', 'SKIPPED', 'Already processed');
      return;
    }
    if (!id || !client) throw new Error('ID and Client are required');
    if (!isEmail_(email)) throw new Error('Invalid e-mail: ' + email);

    const settings = getSettings_();
    const createdNow = !folderUrl;
    if (createdNow) {
      const parent = DriveApp.getFolderById(settings.PARENT_FOLDER_ID);
      const folder = parent.createFolder(safeName_(id + ' - ' + client));
      folder.createFile(
        'request.txt',
        'Request ID: ' + id + '\nClient: ' + client + '\nE-mail: ' + email +
        '\nRequest: ' + request + '\nApproved: ' + new Date().toISOString() + '\n'
      );
      folderUrl = folder.getUrl();
      sh.getRange(row, COL.FOLDER).setValue(folderUrl);
    }

    notify_(settings, { id: id, client: client, email: email, request: request, url: folderUrl });
    sh.getRange(row, COL.NOTIFIED).setValue(new Date());

    log_(id, 'process', 'OK',
      (createdNow ? 'Folder created and notification sent' : 'Notification sent (folder already existed)') +
      ' (' + settings.NOTIFY_MODE + ')');
  } catch (err) {
    log_(id, 'process', 'ERROR', err.message);
  } finally {
    lock.releaseLock();
  }
}

function notify_(settings, d) {
  if (MailApp.getRemainingDailyQuota() < 1) throw new Error('Daily e-mail quota reached');
  const live = settings.NOTIFY_MODE === 'live';
  const to = live ? d.email : settings.OWNER_EMAIL;
  const subject = (live ? '' : '[TEST] ') + 'Request ' + d.id + ' approved';
  const html =
    '<p>Hello ' + esc_(d.client) + ',</p>' +
    '<p>Your request <b>' + esc_(d.id) + '</b> has been approved:</p>' +
    '<blockquote>' + esc_(d.request) + '</blockquote>' +
    '<p>Shared folder: <a href="' + esc_(d.url) + '">open in Google Drive</a></p>' +
    (live ? '' : '<p><i>Test mode: in live mode this e-mail goes to ' + esc_(d.email) + '.</i></p>');
  const msg = { to: to, subject: subject, htmlBody: html, name: 'Requests Desk' };
  if (live) msg.cc = settings.OWNER_EMAIL;
  MailApp.sendEmail(msg);
}

function sendDailySummary() {
  const settings = getSettings_();
  const sh = SpreadsheetApp.getActive().getSheetByName(REQ_SHEET);
  const last = sh.getLastRow();
  const counts = {};
  if (last >= 2) {
    sh.getRange(2, COL.STATUS, last - 1, 1).getValues().forEach(function (r) {
      const s = String(r[0]).trim() || '(blank)';
      counts[s] = (counts[s] || 0) + 1;
    });
  }
  const lines = Object.keys(counts).map(function (k) {
    return '<li>' + esc_(k) + ': ' + counts[k] + '</li>';
  }).join('');
  MailApp.sendEmail({
    to: settings.OWNER_EMAIL,
    subject: 'Daily requests summary',
    htmlBody: '<p>Requests by status:</p><ul>' + (lines || '<li>No requests</li>') + '</ul>' +
      '<p><a href="' + SpreadsheetApp.getActive().getUrl() + '">Open the spreadsheet</a></p>',
    name: 'Requests Desk'
  });
  log_('-', 'daily summary', 'OK', JSON.stringify(counts));
}

/* ---------- Helpers ---------- */

function getSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG_SHEET);
  const out = {};
  const last = sh.getLastRow();
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
    if (r[0]) out[String(r[0]).trim()] = String(r[1]).trim();
  });
  return out;
}

function setSetting_(key, value) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG_SHEET);
  const last = sh.getLastRow();
  const keys = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0]).trim() === key) {
      sh.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

function log_(id, action, result, detail) {
  SpreadsheetApp.getActive().getSheetByName(LOG_SHEET)
    .appendRow([new Date(), id, action, result, detail]);
}

function isEmail_(s) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
}

function safeName_(s) {
  return s.replace(/[\\/:*?"<>|]/g, '-').slice(0, 100);
}

function esc_(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
