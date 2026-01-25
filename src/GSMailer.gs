/************************************
 * GSMailer – FINAL PRODUCTION VERSION
 ************************************/

const SYSTEM_SHEETS = ['SETTINGS', 'EMAIL_RULES'];
const LIB_VERSION = '2.0.3';
const LIB_AUTHOR = 'Rohit Verma';

const SHEETS = {
  SETTINGS: 'SETTINGS',
  RULES: 'EMAIL_RULES'
};

/* =====================
   MENU
===================== */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const mainMenu = ui.createMenu('📬 GSMailer');

  // --- Core actions (most used) ---
  mainMenu
    .addItem('🚀 Send Emails', 'sendEmailsLive')
    .addItem('👁 Preview Email (Selected Row)', 'previewSelectedRow')
    .addSeparator()
    .addItem('🔢 Generate Registration Numbers', 'generateRegistrationNumbers');

  // --- Admin / Setup submenu ---
  const adminMenu = ui.createMenu('⚙️ Admin / Setup')
    .addItem('🧩 Initial Setup (Create / Detect Sheets)', 'initialSetup')
    .addItem('✅ Validate Settings', 'validateSetup')
    .addItem('🧪 Test Emails (Dry Run)', 'sendEmailsDryRun')
    .addItem('ℹ️ About GSMailer', 'showAboutDialog');
    // .addItem('♻️ Reset Email Status (All Rows)', 'resetSentFlags');

  mainMenu
    .addSeparator()
    .addSubMenu(adminMenu)
    .addToUi();
}

function showAboutDialog() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:sans-serif; padding:15px;">
      <h2>GSMailer 📬</h2>
      <p><b>Version:</b> ${LIB_VERSION}</p>
      <p><b>Author:</b> ${LIB_AUTHOR}</p>
      <p>Rule-based event email automation for Google Sheets.</p>
      <p>Visit the GitHub repository for full documentation. https://github.com/rht5/gsmailer</p>
    </div>`
  )
  .setWidth(350)
  .setHeight(240);

  ui.showModalDialog(html, 'About SheetEventFlow');
}



/* =====================
   INITIAL SETUP
===================== */
function initialSetup() {
  const ss = SpreadsheetApp.getActive();

  const userSheets = ss.getSheets()
    .map(s => s.getName())
    .filter(n => !SYSTEM_SHEETS.includes(n));

  const detectedDataSheet = userSheets[0] || 'DATA';

  if (!ss.getSheetByName(SHEETS.SETTINGS)) {
    const s = ss.insertSheet(SHEETS.SETTINGS);
    
    s.getRange('A1:C13').setValues([
      ['SETTINGS (Change value below as needed. Do not change Key)', '', '',],
      ['Key', 'Value', 'Description'],
      ['DATA_SHEET', detectedDataSheet, 'Main data sheet containing attendee rows'],
      ['STATUS_COLUMN', 'Status', 'Column that controls which email rule applies'],
      ['EMAIL_COLUMN', 'Email', 'Recipient email column'],
      ['SENT_FLAG_COLUMN', 'Email Sent', 'System column to mark email as sent'],
      ['LAST_SENT_COLUMN', 'Last Email At', 'Timestamp of email sending'],
      ['ERROR_COLUMN', 'Email Error', 'Error message if sending fails'],
      ['BATCH_LIMIT', '10', 'Max emails per execution'],
      ['REG_NUMBER_ENABLED', 'no', 'Enable auto registration number generation (yes/no)'],
      ['REG_NUMBER_COLUMN', 'Registration No', 'Column to store registration number'],
      ['REG_NUMBER_PATTERN', 'EVT-{{NUMBER}}', 'Optional pattern for registration number'],
      ['REG_NUMBER_START', '001', 'Starting number if column is empty']
    ]);

    s.getRange(1, 1, 1, 3).merge();
  
    s.getRange('E1:F6').setValues([
    ['EVENT_CONTEXT', ''],
    ['Key', 'Value'],
    ['Event Name', 'Self Empowerment'],
    ['Event Date', '18 Jan 2026, 5 PM IST'],
    ['Event Venue', 'Location Here'],
    ['Organizer Name', 'XYZ Foundation'],
    ]);
    s.getRange('E1:F1').merge();

    styleTable(s);
    styleContextTable(s);

    const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['yes', 'no'], true) // true = dropdown
    .setAllowInvalid(false)
    .build();

    s.getRange('B10').setDataValidation(rule);

    s.getRange('B10')
    .setHorizontalAlignment('center')
    .setBackground('#ecfeff'); // subtle highlight

  }


  if (!ss.getSheetByName(SHEETS.RULES)) {
    const sr = ss.insertSheet(SHEETS.RULES);
    sr.getRange('A1:F2').setValues([
      ['Status', 'Subject', 'Template Doc URL', 'CC', 'BCC', 'Attachment URLs'],
      [
        'Approved',
        'Registration Approved for {{Event Name}}',
        'https://docs.google.com/document/d/1Pm04gdIcz-YP3gjV-x1mPT5Zf7lA0MR4ysarJe2tBg4/edit?usp=sharing',
        '',
        '',
        ''
      ]
    ]);


    styleRulesTable(sr);
  }


  //SpreadsheetApp.getUi().alert('Initial setup completed.');
}


function styleTable(s) {
  const tableRange = s.getRange('A1:C13');

  s.getRange('A1:C1')
    .setBackground('#b1c2c7')
    .setFontColor('#000000')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  s.getRange('A2:C2').setFontWeight('bold');
  s.getRange('A2:A13').setBackground('#efefef');

  applyTableBorders(tableRange);

  s.setColumnWidths(1, 1, 200);
  s.setColumnWidths(2, 1, 130);
  s.setColumnWidths(3, 1, 320);
}


function styleContextTable(s) {
  const tableRange = s.getRange('E1:F13');

  s.getRange('E1:F1')
    .setBackground('#e4cfce')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  s.getRange('E2:F2').setFontWeight('bold');
  s.getRange('E3:F13').setBackground('#f9fafb');

  applyTableBorders(tableRange);

  s.setColumnWidths(5, 1, 120);
  s.setColumnWidths(6, 1, 180);
}


function styleRulesTable(s) {
  const tableRange = s.getRange('A1:F13');

  s.getRange('A1:F1')
    .setBackground('#b1c2c7')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  s.getRange('A2:F13').setBackground('#f9fafb');

  applyTableBorders(tableRange);

  s.setColumnWidths(1, 1, 110);
  s.setColumnWidths(2, 1, 230);
  s.setColumnWidths(3, 1, 320);
  s.setColumnWidths(6, 1, 120);

  s.getRange('B2:C2').setWrap(true);
}

function applyTableBorders(range) {
  range.setBorder(
    true, true, true, true,
    true, true,
    '#d1d5db',
    SpreadsheetApp.BorderStyle.SOLID
  );
}


/* =====================
   ENTRY POINTS
===================== */
function sendEmailsDryRun() {
  processEmails(true);
}

function sendEmailsLive() {
  processEmails(false);
}

/* =====================
   REGISTRATION NUMBER
===================== */
function generateRegistrationNumbers() {
  const settings = loadSettings();
  Logger.log(settings.REG_NUMBER_ENABLED);
  if (settings.REG_NUMBER_ENABLED !== 'yes') {
    return alertUser('Registration number feature is disabled.');
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName(settings.DATA_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const map = mapHeaders(headers);

  const col = settings.REG_NUMBER_COLUMN;
  if (!(col in map)) throw new Error(`Missing column: ${col}`);

  let lastNumber = Number(settings.REG_NUMBER_START || 1) - 1;

  for (let i = 1; i < data.length; i++) {
    const cell = data[i][map[col]];
    if (cell) {
      const n = Number(String(cell).match(/\d+/));
      if (!isNaN(n)) lastNumber = n;
      continue;
    }

    lastNumber++;
    const value = settings.REG_NUMBER_PATTERN
      ? settings.REG_NUMBER_PATTERN.replace('{{NUMBER}}', lastNumber)
      : lastNumber;

    sheet.getRange(i + 1, map[col] + 1).setValue(value);
  }

  //alertUser('Registration numbers generated.');
}

/* =====================
   CORE PROCESSOR 
===================== */
function processEmails(isDryRun) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another process is running.');

  try {
    const ss = SpreadsheetApp.getActive();
    const settings = loadSettings();
    const rules = loadRules();
    const eventContext = loadEventContext();

    const sheet = ss.getSheetByName(settings.DATA_SHEET);
    if (!sheet) throw new Error('DATA sheet not found');

    const values = sheet.getDataRange().getValues();
    const headers = values.shift();
    const map = mapHeaders(headers);

    validateRequiredColumns(settings, map);
    
    const dataHeaders = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];

    

    validateTemplateVariables(
      rules,
      dataHeaders,
      eventContext
    );

    const limit = Number(settings.BATCH_LIMIT || 50);
    let count = 0;

    for (let i = 0; i < values.length && count < limit; i++) {
      const row = values[i];
      const rowIndex = i + 2;

      try {
        const status = row[map[settings.STATUS_COLUMN]];
        const email = row[map[settings.EMAIL_COLUMN]];
        const sent = row[map[settings.SENT_FLAG_COLUMN]];

        if (!email || sent || !rules[status]) continue;


        const ctx = {
          ...eventContext,
          ...buildContext(headers, row)
        };

        const subject = renderTemplate(rules[status].subject, ctx);

        const bodyHtml = renderTemplate(rules[status].body, ctx);

        if (!isDryRun) {

            const attachments = getAttachmentBlobsFromUrls(
              rules[status].attachmentUrls
            );

            GmailApp.sendEmail(
              email,
              subject,
              "Please view this email in an HTML compatible client.",
              {
                htmlBody: bodyHtml,
                cc: rules[status].cc || '',
                bcc: rules[status].bcc || '',
                attachments: attachments
              }
            );


          sheet.getRange(rowIndex, map[settings.SENT_FLAG_COLUMN] + 1).setValue('YES');
          sheet.getRange(rowIndex, map[settings.LAST_SENT_COLUMN] + 1).setValue(new Date());
        } else {
          sheet.getRange(rowIndex, map[settings.SENT_FLAG_COLUMN] + 1).setValue('Test - Success');
          sheet.getRange(rowIndex, map[settings.LAST_SENT_COLUMN] + 1).setValue(new Date());
        }

        //logEvent(rowIndex, email, status, isDryRun ? 'DRY RUN' : 'SENT');
        count++;

      } catch (e) {
        sheet.getRange(rowIndex, map[settings.SENT_FLAG_COLUMN] + 1).setValue('FAILED');
        sheet.getRange(rowIndex, map[settings.ERROR_COLUMN] + 1).setValue(e.message);
        //logEvent(rowIndex, '', '', 'ERROR: ' + e.message);
      }
    }

    alertUser(`Completed. Rows processed: ${count}`);
  } finally {
    lock.releaseLock();
  }
}

/* =====================
   PREVIEW
===================== */
function previewSelectedRow() {
  const settings = loadSettings();
  const rules = loadRules();
  const eventContext = loadEventContext();

  const sheet = SpreadsheetApp.getActive().getSheetByName(settings.DATA_SHEET);
  const rowIndex = sheet.getActiveRange().getRow();
  if (rowIndex < 2) return alertUser('Select a valid row.');

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const row = data[rowIndex - 1];
  const map = mapHeaders(headers);

  const status = row[map[settings.STATUS_COLUMN]];
  if (!rules[status]) return alertUser('No rule for this status.');

  const ctx = { ...eventContext, ...buildContext(headers, row) };

  const subject = renderTemplate(rules[status].subject, ctx);
  const bodyHtml = renderTemplate(rules[status].body, ctx);

  // Removed <pre> tags and added simple styling wrapper
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family: sans-serif; padding: 10px;">
       <h3 style="border-bottom: 1px solid #ccc; padding-bottom: 10px;">Subject: ${subject}</h3>
       <div>${bodyHtml}</div>
     </div>`
  ).setWidth(650).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Email Preview');
}

/* =====================
   LOADERS / VALIDATION / UTIL
===================== */

function loadSettings() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.SETTINGS);

  // Get only columns A to C, starting from row 3 (skip header)
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(3, 1, lastRow - 1, 3).getValues();

  const s = {};
  data.forEach(([k, v, _c]) => {
    s[k] = v; // column C is ignored for now
  });

  return s;
}



function loadRules() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.RULES);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const map = mapHeaders(headers);
  const rules = {};
  data.forEach(r => {
    rules[r[map.Status]] = {
      subject: r[map.Subject],
      body: getDocBody(r[map['Template Doc URL']]),
      cc: r[map.CC],
      bcc: r[map.BCC],
      attachmentUrls: r[map['Attachment URLs']] || ''
    };
  });
  return rules;
}

function loadEventContext() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.SETTINGS);
  if (!sheet) return {};

  const lastRowEF = Math.max(
    sheet.getRange("E:E").getLastRow(),
    sheet.getRange("F:F").getLastRow()
  );

  if (lastRowEF < 2) return {};

  const data = sheet.getRange(3, 5, lastRowEF - 1, 2).getValues(); // E:F
  const ctx = {};

  data.forEach(([k, v]) => {
    if (k) ctx[k] = v ?? '';
  });

  return ctx;
}



function validateTemplateVariables(rules, dataHeaders, eventContext) {
  // Schema only — no row data
  const declared = new Set([
    ...dataHeaders,
    ...Object.keys(eventContext)
  ]);

  const missing = new Set();

  Object.values(rules).forEach(rule => {
    extractVars(rule.subject).forEach(v => {
      if (!declared.has(v)) missing.add(v);
    });

    if (rule.body) {
      extractVars(rule.body).forEach(v => {
        if (!declared.has(v)) missing.add(v);
      });
    }
  });

  if (missing.size > 0) {
    throw new Error(
      'Missing template variables: ' + [...missing].join(', ')
    );
  }
}


function extractVars(text) {
  if (!text) return [];
  const vars = new Set();
  const re = /{{\s*([^}]+)\s*}}/g;
  let m;
  while ((m = re.exec(text))) vars.add(m[1]);
  return [...vars];
}


function renderTemplate(t, ctx) {
  return t.replace(/{{\s*([^}]+)\s*}}/g, (_, k) => ctx[k] ?? '');
}

function validateRequiredColumns(s, map) {
  [s.STATUS_COLUMN, s.EMAIL_COLUMN, s.SENT_FLAG_COLUMN]
    .forEach(c => !(c in map) && (() => { throw new Error(`Missing column: ${c}`); })());
}

function mapHeaders(h) {
  const m = {};
  h.forEach((x, i) => m[x] = i);
  return m;
}

function buildContext(h, r) {
  const c = {};
  h.forEach((x, i) => c[x] = r[i] ?? '');
  return c;
}

function getDocBody(url) {
  DriveApp.getRootFolder(); // ensure Drive scope
  const id = url.match(/[-\w]{25,}/)[0];

  const exportUrl =
    "https://www.googleapis.com/drive/v3/files/" +
    id +
    "/export?mimeType=text/html";

  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Failed to fetch Doc as HTML');
  }

  const html = response.getContentText();

  // --- Extract <body> content only ---
  const bodyMatch = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  let body = bodyMatch ? bodyMatch[1] : html;

  // --- HARD RESET GOOGLE DOC MARGINS ---
  body = `
    <div style="
      margin:0;
      padding:0;
      font-family: Arial, Helvetica, sans-serif;
      font-size:14px;
      line-height:1.5;
    ">
      ${body}
    </div>
  `;

  return body;
}


function resetSentFlags() {
  const s = loadSettings();
  const sheet = SpreadsheetApp.getActive().getSheetByName(s.DATA_SHEET);
  const data = sheet.getDataRange().getValues();
  const map = mapHeaders(data[0]);

  for (let i = 1; i < data.length; i++) {
    sheet.getRange(i + 1, map[s.SENT_FLAG_COLUMN] + 1).clearContent();
    sheet.getRange(i + 1, map[s.LAST_SENT_COLUMN] + 1).clearContent();
    sheet.getRange(i + 1, map[s.ERROR_COLUMN] + 1).clearContent();
  }
  alertUser('Email flags reset.');
}



const SETTINGS_SCHEMA = {
  DATA_SHEET:        { required: true },
  STATUS_COLUMN:     { required: true },
  EMAIL_COLUMN:      { required: true },
  SENT_FLAG_COLUMN:  { required: true },
  LAST_SENT_COLUMN:  { required: false },
  ERROR_COLUMN:      { required: false },
  BATCH_LIMIT:       { required: true, type: 'number' },

  REG_NUMBER_ENABLED:{ required: false, type: 'boolean' },
  REG_NUMBER_COLUMN: { requiredIf: 'REG_NUMBER_ENABLED' },
  REG_NUMBER_START:  { requiredIf: 'REG_NUMBER_ENABLED', type: 'number' },
  REG_NUMBER_PATTERN:{ required: false }
};


function validateSetup() {

  const settings = loadSettings();
  const rules = loadRules();
  const ctx = loadEventContext();

  validateSettings(settings);
  validateDataSheet(settings);
  validateRules(rules);
  validateTemplateVariables(
    rules,
    getDataHeaders(settings),
    ctx
  );

  alertUser('Setup validated successfully.');
}


function validateSettings(settings) {
  const errors = [];

  Object.entries(SETTINGS_SCHEMA).forEach(([key, rule]) => {
    const value = settings[key];

    // Required
    if (rule.required && !value) {
      errors.push(`Missing value for setting: ${key}`);
    }

    // Required if enabled
    if (rule.requiredIf) {
      const enabled = isTruthy(settings[rule.requiredIf]);
      if (enabled && !value) {
        errors.push(`${key} is required when ${rule.requiredIf} is enabled`);
      }
    }

    // Type checks
    if (value && rule.type === 'number' && isNaN(Number(value))) {
      errors.push(`${key} must be a number`);
    }

    if (value && rule.type === 'boolean' && !isBooleanLike(value)) {
      errors.push(`${key} must be TRUE or FALSE`);
    }
  });

  if (errors.length) {
    throw new Error('Settings validation failed:\n• ' + errors.join('\n• '));
  }
}




function alertUser(m) {
  SpreadsheetApp.getUi().alert(m);
}

function extractDriveFileId(url) {
  if (!url) return null;
  const m = url.match(/[-\w]{25,}/);
  return m ? m[0] : null;
}

function getAttachmentBlobsFromUrls(urlString) {
  if (!urlString) return [];

  const urls = urlString
    .split(',')
    .map(u => u.trim())
    .filter(Boolean);

  const blobs = [];

  urls.forEach(url => {
    const id = extractDriveFileId(url);
    if (!id) {
      throw new Error('Invalid attachment URL: ' + url);
    }
    try {
      blobs.push(DriveApp.getFileById(id).getBlob());
    } catch (e) {
      throw new Error('Cannot access attachment: ' + url);
    }
  });

  return blobs;
}

function validateDataSheet(settings) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(settings.DATA_SHEET);

  if (!sheet) {
    throw new Error(`DATA_SHEET "${settings.DATA_SHEET}" not found`);
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = mapHeaders(headers);

  [
    settings.STATUS_COLUMN,
    settings.EMAIL_COLUMN,
    settings.SENT_FLAG_COLUMN
  ].forEach(col => {
    if (col && !(col in map)) {
      throw new Error(`Missing column in data sheet: ${col}`);
    }
  });
}

function validateRules(rules) {
  if (!Object.keys(rules).length) {
    throw new Error('No email rules found in EMAIL_RULES sheet');
  }

  Object.entries(rules).forEach(([status, rule]) => {
    if (!status) throw new Error('EMAIL_RULES has empty Status value');
    if (!rule.subject) throw new Error(`Missing subject for status: ${status}`);
    if (!rule.body) throw new Error(`Missing template for status: ${status}`);
  });
}

function isTruthy(v) {
  return ['TRUE', 'YES', '1'].includes(String(v).toUpperCase());
}

function isBooleanLike(v) {
  return ['TRUE', 'FALSE', 'YES', 'NO', '1', '0'].includes(String(v).toUpperCase());
}

function getDataHeaders(settings) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(settings.DATA_SHEET);
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}


var GSMailer = {
  onOpen,
  initialSetup,
  validateSetup,
  sendEmailsLive,
  sendEmailsDryRun,
  previewSelectedRow,
  generateRegistrationNumbers
};
