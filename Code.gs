/**
 * SMART MAINTAIN – Bharat Petroleum, Gonda LPG Plant
 * Google Apps Script backend for shared equipment/AMC data and automatic email reminders.
 *
 * IMPORTANT ACCESS NOTE:
 * This version accepts website sync writes without an OTP because the current website
 * has no viewer/editor login. Therefore the Web App endpoint must be restricted to
 * "Anyone within Bharat Petroleum" if available. If deployed as public "Anyone",
 * anyone who obtains the endpoint can change the shared register. Obtain BPCL IT approval.
 *
 * INSTALL:
 * 1. Replace Code.gs in the EXISTING Apps Script project connected to the website URL.
 * 2. Save.
 * 3. Run setupSmartMaintain() once and approve requested Google permissions.
 * 4. Run sendTestEmail() once and approve MailApp permission. It sends a real test email
 *    to all five configured recipients.
 * 5. Deploy > Manage deployments > Edit > New version > Deploy.
 * 6. Keep Execute as: Me. Keep the same /exec URL in index.html.
 * 7. Project Settings > Time zone: Asia/Kolkata.
 */

const RECIPIENTS = [
  'rithikeshrathi@bharatpetroleum.in',
  'gauravmadan@bharatpetroleum.in',
  'patilvishalravi@bharatpetroleum.in',
  'atharavsanjayka@bharatpetroleum.in',
  'frootyrithikesh@gmail.com'
];

const REMINDER_DAYS_DEFAULT = [60, 45, 30, 15, 10, 5, 2, 0];
const TZ = 'Asia/Kolkata';
const PROP_KEY = 'SMART_MAINTAIN_SPREADSHEET_ID';
const REMINDER_DAYS_KEY = 'SMART_MAINTAIN_REMINDER_DAYS';
const REMINDERS_ENABLED_KEY = 'SMART_MAINTAIN_REMINDERS_ENABLED';

function setupSmartMaintain() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty(PROP_KEY);
  let ss;

  if (spreadsheetId) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } else {
    ss = SpreadsheetApp.create('SMART MAINTAIN – Gonda LPG Plant AMC Register');
    props.setProperty(PROP_KEY, ss.getId());
  }

  ensureSheet_(ss, 'Equipment Register', ['Equipment ID', 'Equipment JSON']);
  ensureSheet_(ss, 'AMC Contracts', ['AMC ID', 'Contract JSON']);
  ensureSheet_(ss, 'Maintenance Logs', ['Log ID', 'Log JSON']);
  ensureSheet_(ss, 'Email Log', [
    'Sent At (IST)', 'AMC ID', 'Contract Ref', 'Expiry Date', 'Reminder Days', 'Recipients'
  ]);

  // Use the requested reminder pattern as the initial default.
  if (!props.getProperty(REMINDER_DAYS_KEY)) {
    props.setProperty(REMINDER_DAYS_KEY, JSON.stringify(REMINDER_DAYS_DEFAULT));
  }
  if (!props.getProperty(REMINDERS_ENABLED_KEY)) {
    props.setProperty(REMINDERS_ENABLED_KEY, 'true');
  }

  installDailyReminderTrigger_();

  Logger.log('SMART MAINTAIN database: ' + ss.getUrl());
  Logger.log('Daily AMC reminder trigger installed. Time zone: ' + TZ);
  Logger.log('Reminder recipients: ' + RECIPIENTS.join(', '));
  Logger.log('Existing spreadsheet data has been retained.');
  return ss.getUrl();
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  return sheet;
}

function installDailyReminderTrigger_() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkAmcExpiry');

  // Keep exactly one daily trigger.
  triggers.slice(1).forEach(t => ScriptApp.deleteTrigger(t));
  if (triggers.length === 0) {
    ScriptApp.newTrigger('checkAmcExpiry')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();
  }
}

function reinstallDailyReminderTrigger() {
  installDailyReminderTrigger_();
  Logger.log('Daily AMC reminder trigger is installed for approximately 8 AM in the script time zone.');
}

function getDb_() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP_KEY);
  if (!id) throw new Error('Database is not initialized. Run setupSmartMaintain() first.');
  return SpreadsheetApp.openById(id);
}

function jsonResponse_(result, callback) {
  const json = JSON.stringify(result);
  if (callback && /^[a-zA-Z_$][0-9a-zA-Z_$\.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  const callback = params.callback || '';
  const action = params.action || '';

  try {
    if (action === 'get') {
      const ss = getDb_();
      return jsonResponse_({
        ok: true,
        assets: readJsonRows_(ss.getSheetByName('Equipment Register')),
        contracts: readJsonRows_(ss.getSheetByName('AMC Contracts')),
        logs: readJsonRows_(ss.getSheetByName('Maintenance Logs'))
      }, callback);
    }

    if (action === 'health') {
      const ss = getDb_();
      const triggers = ScriptApp.getProjectTriggers()
        .filter(t => t.getHandlerFunction() === 'checkAmcExpiry');
      return jsonResponse_({
        ok: true,
        databaseReady: true,
        triggerInstalled: triggers.length > 0,
        triggerCount: triggers.length,
        timeZone: Session.getScriptTimeZone(),
        contractCount: readJsonRows_(ss.getSheetByName('AMC Contracts')).length,
        reminderDays: getReminderDays_(),
        remindersEnabled: remindersEnabled_()
      }, callback);
    }

    throw new Error('Unsupported action.');
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err.message || err) }, callback);
  }
}

/**
 * The website posts JSON as text/plain with mode:no-cors.
 * This endpoint intentionally accepts sync without a login token to match the current
 * open-edit website. Restrict the Apps Script deployment to BPCL users where possible.
 */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action !== 'sync') throw new Error('Unsupported action.');

    const ss = getDb_();
    ensureSheet_(ss, 'Equipment Register', ['Equipment ID', 'Equipment JSON']);
    ensureSheet_(ss, 'AMC Contracts', ['AMC ID', 'Contract JSON']);
    ensureSheet_(ss, 'Maintenance Logs', ['Log ID', 'Log JSON']);
    ensureSheet_(ss, 'Email Log', [
      'Sent At (IST)', 'AMC ID', 'Contract Ref', 'Expiry Date', 'Reminder Days', 'Recipients'
    ]);

    writeJsonRows_(ss.getSheetByName('Equipment Register'), body.assets || [], 'id');
    writeJsonRows_(ss.getSheetByName('AMC Contracts'), body.contracts || [], 'id');
    writeJsonRows_(ss.getSheetByName('Maintenance Logs'), body.logs || [], 'id');

    if (body.reminderPrefs) {
      if (body.reminderPrefs.days !== undefined) {
        const parsedDays = parseReminderDays_(body.reminderPrefs.days);
        PropertiesService.getScriptProperties().setProperty(REMINDER_DAYS_KEY, JSON.stringify(parsedDays));
      }
      if (body.reminderPrefs.enabled !== undefined) {
        PropertiesService.getScriptProperties().setProperty(
          REMINDERS_ENABLED_KEY,
          body.reminderPrefs.enabled === true || String(body.reminderPrefs.enabled) === 'true' ? 'true' : 'false'
        );
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      ok: true,
      syncedAt: new Date().toISOString(),
      assets: (body.assets || []).length,
      contracts: (body.contracts || []).length,
      logs: (body.logs || []).length
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error('SMART MAINTAIN sync failed: ' + String(err.stack || err));
    return ContentService.createTextOutput(JSON.stringify({
      ok: false,
      error: String(err.message || err)
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function readJsonRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues()
    .map(row => {
      try { return JSON.parse(row[0]); } catch (_) { return null; }
    })
    .filter(Boolean);
}

function writeJsonRows_(sheet, items, key) {
  const safeItems = Array.isArray(items) ? items : [];
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  }
  if (!safeItems.length) return;
  const rows = safeItems.map(item => [String(item[key] || ''), JSON.stringify(item)]);
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

function parseReminderDays_(value) {
  const values = Array.isArray(value) ? value : String(value || '').split(',');
  const days = values.map(x => Number(String(x).trim()));
  if (!days.length || days.some(x => !Number.isInteger(x) || x < 0 || x > 3650)) {
    throw new Error('Reminder days must be whole numbers from 0 to 3650.');
  }
  return [...new Set(days)].sort((a, b) => b - a);
}

function getReminderDays_() {
  const saved = PropertiesService.getScriptProperties().getProperty(REMINDER_DAYS_KEY);
  if (!saved) return REMINDER_DAYS_DEFAULT;
  try { return parseReminderDays_(JSON.parse(saved)); }
  catch (_) { return REMINDER_DAYS_DEFAULT; }
}

function remindersEnabled_() {
  return PropertiesService.getScriptProperties().getProperty(REMINDERS_ENABLED_KEY) !== 'false';
}

function parseExpiryDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  }
  const raw = String(value || '').trim();
  if (!raw) return null;

  // Preferred website format: yyyy-MM-dd.
  let m = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[3]).padStart(2, '0');

  // Also accept dd-MM-yyyy or dd/MM/yyyy from manually imported legacy records.
  m = raw.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return m[3] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[1]).padStart(2, '0');

  return null;
}

function checkAmcExpiry() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    Logger.log('Another AMC reminder check is already running. This run was skipped.');
    return;
  }

  try {
    if (!remindersEnabled_()) {
      Logger.log('AMC email reminders are disabled in website preferences.');
      return;
    }

    const ss = getDb_();
    const contracts = readJsonRows_(ss.getSheetByName('AMC Contracts'));
    const logs = readJsonRows_(ss.getSheetByName('Maintenance Logs'));
    const emailLog = ensureSheet_(ss, 'Email Log', [
      'Sent At (IST)', 'AMC ID', 'Contract Ref', 'Expiry Date', 'Reminder Days', 'Recipients'
    ]);

    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    const todayUtc = Date.parse(today + 'T00:00:00Z');
    const reminderDays = getReminderDays_();
    const alreadySent = new Set();

    if (emailLog.getLastRow() > 1) {
      emailLog.getRange(2, 2, emailLog.getLastRow() - 1, 4).getValues().forEach(row => {
        // Columns B-E: AMC ID, Contract Ref, Expiry Date, Reminder Days.
        alreadySent.add([row[0], row[2], row[3]].join('|'));
      });
    }

    const due = [];
    contracts.forEach(c => {
      if (!c || !c.id) return;
      const expiry = parseExpiryDate_(c.end);
      if (!expiry) return;

      const days = Math.round((Date.parse(expiry + 'T00:00:00Z') - todayUtc) / 86400000);

      // Before expiry: only configured reminder dates.
      // On expiry day (0) and every day after expiry (-1, -2, -3 ...): send daily.
      if (days > 0 && !reminderDays.includes(days)) return;

      const dedupeKey = [c.id, expiry, days].join('|');
      if (alreadySent.has(dedupeKey)) return;

      const visitsGiven = logs.filter(l =>
        l && l.type === 'AMC Visit' && String(l.contract || '') === String(c.id)
      ).length;

      const agreed = c.totalVisits === '' || c.totalVisits == null ? '' : Number(c.totalVisits);
      const remaining = agreed === '' ? '' : Math.max(0, agreed - visitsGiven);

      due.push({
        id: c.id,
        ref: c.ref || c.id,
        equipment: c.equipment || '',
        oem: c.oem || c.agency || '',
        start: parseExpiryDate_(c.start) || '',
        expiry: expiry,
        days: days,
        po: c.po || '',
        ola: c.ola || '',
        totalVisits: agreed,
        visitsGiven: visitsGiven,
        visitsRemaining: remaining
      });
    });

    if (!due.length) {
      Logger.log('No AMC reminders due today (' + today + '). Contracts checked: ' + contracts.length);
      return;
    }

    const subject = 'SMART MAINTAIN | AMC expiry reminder | Gonda LPG Plant';
    const lines = [
      'Dear Team Gonda,',
      '',
      'The following AMC contract(s) have reached a configured reminder date.',
      'Reminder dates are calculated from the exact validity end date entered in the contract.',
      ''
    ];

    due.forEach((x, i) => {
      lines.push((i + 1) + '. Equipment: ' + x.equipment);
      lines.push('   Contract reference: ' + x.ref);
      lines.push('   OEM / Agency: ' + x.oem);
      lines.push('   Validity start: ' + x.start);
      lines.push('   Exact validity end: ' + x.expiry);
      lines.push('   Days to expiry: ' + x.days);
      if (x.po) lines.push('   PO: ' + x.po);
      if (x.ola) lines.push('   OLA: ' + x.ola);
      if (x.totalVisits !== '') {
        lines.push('   Visits agreed / given / remaining: ' +
          x.totalVisits + ' / ' + x.visitsGiven + ' / ' + x.visitsRemaining);
      }
      lines.push('');
    });

    lines.push(
      'Please review the contract and arrange the required action.',
      '',
      'Regards,',
      'SMART MAINTAIN – Gonda LPG Plant'
    );

    const remainingQuota = MailApp.getRemainingDailyQuota();
    if (remainingQuota < RECIPIENTS.length) {
      throw new Error('Mail quota is insufficient. Remaining recipient quota: ' + remainingQuota);
    }

    MailApp.sendEmail({
      to: RECIPIENTS.join(','),
      subject: subject,
      body: lines.join('\n'),
      name: 'SMART MAINTAIN – Gonda LPG Plant'
    });

    const sentAt = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    const rows = due.map(x => [sentAt, x.id, x.ref, x.expiry, x.days, RECIPIENTS.join(', ')]);
    emailLog.getRange(emailLog.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    Logger.log('SUCCESS: sent one AMC reminder email to all five recipients for ' +
      due.length + ' contract(s).');
  } catch (err) {
    console.error('AMC reminder email failed: ' + String(err.stack || err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/** Sends a test email to verify SMART MAINTAIN web application email notifications. */
function sendTestEmail() {
  const quota = MailApp.getRemainingDailyQuota();

  if (quota < RECIPIENTS.length) {
    throw new Error('Mail quota is insufficient. Remaining recipient quota: ' + quota);
  }

  MailApp.sendEmail({
    to: RECIPIENTS.join(','),
    subject: 'SMART MAINTAIN Web Application – Email Notification Test',
    body: [
      'Dear Team,',
      '',
      'This is a test email from the SMART MAINTAIN Web Application – Gonda LPG Plant.',
      '',
      'This email is being sent to verify whether the email notification functionality of the SMART MAINTAIN Web Application is working properly and whether email notifications are being successfully delivered to the configured recipients.',
      '',
      'You are requested to kindly ignore this email. No action is required.',
      '',
      'Kindly ignore this email if you are not an intended recipient.',
      '',
      'Regards,',
      'SMART MAINTAIN',
      'Gonda LPG Plant',
      'Bharat Petroleum Corporation Limited'
    ].join('\n'),
    name: 'SMART MAINTAIN – Gonda LPG Plant'
  });

  Logger.log('TEST EMAIL SENT to: ' + RECIPIENTS.join(', '));
}

/** Run manually to inspect the database, trigger, and remaining email quota. */
function diagnoseSmartMaintain() {
  const ss = getDb_();
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkAmcExpiry');
  const contracts = readJsonRows_(ss.getSheetByName('AMC Contracts'));
  Logger.log('Spreadsheet: ' + ss.getUrl());
  Logger.log('Contracts currently stored in cloud: ' + contracts.length);
  Logger.log('AMC reminder trigger count: ' + triggers.length);
  Logger.log('Script time zone: ' + Session.getScriptTimeZone());
  Logger.log('Reminder days: ' + JSON.stringify(getReminderDays_()));
  Logger.log('Reminders enabled: ' + remindersEnabled_());
  Logger.log('Remaining daily email recipient quota: ' + MailApp.getRemainingDailyQuota());
  if (triggers.length === 0) Logger.log('ACTION REQUIRED: run setupSmartMaintain() to install the daily trigger.');
}

/**
 * Run this function while signed in to the dedicated corporate Gmail account.
 * It reports the account Apps Script sees as the effective execution identity.
 * MailApp sends as the account that authorizes/runs the script or its trigger.
 */
function showSenderAccount() {
  const email = Session.getEffectiveUser().getEmail();
  Logger.log('Apps Script effective sender account: ' + (email || '[not exposed by Google in this execution context]'));
  Logger.log('Reminder recipients: ' + RECIPIENTS.join(', '));
  Logger.log('Reminder trigger time zone: ' + Session.getScriptTimeZone());
  Logger.log('Remaining daily recipient quota: ' + MailApp.getRemainingDailyQuota());
  return email || '';
}
