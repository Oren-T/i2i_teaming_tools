// ============================================================
//  HELPERS — Department Task Tracker
// ============================================================
//  Utility functions used by the main entry points in Code.js.
// ============================================================

/**
 * Builds a column-index lookup map from Row 1 headers.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object<string, number>} header string → 0-based column index
 */
function buildColumnMap(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i]).trim();
    if (h) map[h] = i;
  }
  return map;
}

/**
 * Loads the Directory tab into a name → email lookup map.
 * Names are lowercased for case-insensitive matching.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} colMap
 * @returns {Map<string, string>} lowercase name → lowercase email
 */
function loadDirectory(sheet, colMap) {
  const nameIdx = colMap[DIRECTORY_COLUMNS.NAME];
  const emailIdx = colMap[DIRECTORY_COLUMNS.EMAIL];
  const data = sheet.getDataRange().getValues();
  const dir = new Map();

  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][nameIdx] || '').trim();
    const email = String(data[i][emailIdx] || '').trim();
    if (name && email) {
      dir.set(name.toLowerCase(), email.toLowerCase());
    }
  }

  if (dir.size === 0) {
    console.warn('Directory tab is empty — no staff entries found.');
  }
  return dir;
}

/**
 * Loads a project-title → status map from the Projects tab.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} colMap
 * @returns {Map<string, string>} project title → status string
 */
function loadProjectStatusMap(sheet, colMap) {
  const titleIdx = colMap[PROJECT_COLUMNS.TITLE];
  const statusIdx = colMap[PROJECT_COLUMNS.STATUS];
  const data = sheet.getDataRange().getValues();
  const map = new Map();

  for (let i = 1; i < data.length; i++) {
    const title = String(data[i][titleIdx] || '').trim();
    const status = String(data[i][statusIdx] || '').trim();
    if (title) map.set(title, status);
  }
  return map;
}

/**
 * Resolves a comma-separated assignee string to an array of email addresses
 * using the Directory lookup. Skips names not found.
 * @param {string} assigneeStr - Comma-separated display names
 * @param {Map<string, string>} directory - name → email map
 * @returns {string[]}
 */
function resolveAssigneeEmails(assigneeStr, directory) {
  if (!assigneeStr) return [];
  const names = String(assigneeStr).split(',').map(s => s.trim()).filter(Boolean);
  const emails = [];
  for (const name of names) {
    const email = directory.get(name.toLowerCase());
    if (email) {
      emails.push(email);
    }
    else {
      console.warn(`Assignee "${name}" not found in Directory.`);
    }
  }
  return emails;
}

/**
 * Replaces {{TOKEN}} placeholders in a template string.
 * @param {string} template
 * @param {Object<string, string>} tokens - token name (without braces) → value
 * @returns {string}
 */
function substituteTokens(template, tokens) {
  let result = template;
  for (const key of Object.keys(tokens)) {
    result = result.replace(new RegExp(`\\{\\{${key}\\}\\}`, 'g'), () => tokens[key] ?? '');
  }
  return result;
}

/**
 * Escapes special HTML characters so user-entered text renders safely
 * inside HTML email bodies without breaking layout.
 * @param {string} str
 * @returns {string}
 */
function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Formats a Date as YYYY-MM-DD for the Calendar API.
 * @param {Date} date
 * @returns {string}
 */
function formatDateISO(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

/**
 * Formats a Date for human-readable display (e.g. "January 15, 2025").
 * @param {Date} date
 * @returns {string}
 */
function formatDateReadable(date) {
  const months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  return `${months[date.getMonth()]} ${date.getDate()}, ${date.getFullYear()}`;
}

/**
 * Returns the number of whole days from today to targetDate.
 * Positive = future, negative = past.
 * @param {Date} targetDate
 * @returns {number}
 */
function daysUntil(targetDate) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const target = new Date(targetDate);
  target.setHours(0, 0, 0, 0);
  return Math.round((target - today) / 86400000);
}

// ===== Cell Auto-Notes =====

const AUTO_NOTE_PREFIX = '[Auto] ';

/**
 * Sets an automated note on a cell. Leaves user-added notes untouched.
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @param {string} message
 */
function setAutoNote(range, message) {
  const existing = range.getNote();
  if (existing && !existing.startsWith(AUTO_NOTE_PREFIX)) return;
  range.setNote(AUTO_NOTE_PREFIX + message);
}

/**
 * Clears an automated note from a cell. Leaves user-added notes untouched.
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 */
function clearAutoNote(range) {
  if (range.getNote().startsWith(AUTO_NOTE_PREFIX)) {
    range.clearNote();
  }
}

/**
 * Sends an error notification email to all ERROR_EMAILS recipients.
 * @param {string} message
 */
function sendErrorEmail(message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = ss ? ss.getUrl() : '(unknown)';
  const subject = EMAIL_TEMPLATES.ERROR.SUBJECT;
  const body = substituteTokens(EMAIL_TEMPLATES.ERROR.BODY, {
    ERROR_MESSAGE: escapeHtml(message),
    SPREADSHEET_URL: url
  });
  try {
    GmailApp.sendEmail(ERROR_EMAILS.join(','), subject, '', { htmlBody: body });
  }
  catch (e) {
    console.error(`Failed to send error email: ${e.message}`);
  }
}
