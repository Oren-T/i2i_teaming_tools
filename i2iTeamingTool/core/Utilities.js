/**
 * Utility functions for the i2i Teaming Tool library.
 * Contains retry logic, date formatting, token substitution, and other helpers.
 */

/**
 * Executes a function with exponential backoff and jitter for transient errors.
 * Use for Drive, Calendar, Gmail API calls that may hit rate limits.
 * @param {Function} fn - The function to execute (should be idempotent)
 * @param {Object} options - Configuration options
 * @param {number} options.attempts - Max retry attempts (default: 5)
 * @param {number} options.baseMs - Base delay in milliseconds (default: 250)
 * @param {number} options.maxMs - Maximum delay in milliseconds (default: 5000)
 * @returns {*} The result of the function
 * @throws {Error} The last error if all attempts fail
 */
function withBackoff(fn, options = {}) {
  const {
    attempts = 5,
    baseMs = 250,
    maxMs = 5000
  } = options;

  const retryPatterns = [
    /rate limit/i,
    /too many times/i,
    /unavailable/i,
    /429/,
    /5\d\d/,
    /service invoked too many times/i,
    /quota/i
  ];

  let lastError;

  for (let attempt = 1; attempt <= attempts; attempt++) {
    try {
      return fn();
    } catch (error) {
      lastError = error;
      const errorMessage = String(error.message || error);

      const isRetryable = retryPatterns.some(pattern => pattern.test(errorMessage));

      if (!isRetryable || attempt === attempts) {
        DEBUG && console.log(`withBackoff: Non-retryable error or max attempts reached (attempt ${attempt}): ${errorMessage}`);
        throw error;
      }

      // Exponential backoff with full jitter
      const exponentialDelay = Math.min(baseMs * Math.pow(2, attempt - 1), maxMs);
      const jitteredDelay = Math.random() * exponentialDelay;

      DEBUG && console.log(`withBackoff: Attempt ${attempt} failed, retrying in ${Math.round(jitteredDelay)}ms: ${errorMessage}`);
      Utilities.sleep(jitteredDelay);
    }
  }

  throw lastError;
}

/**
 * Formats a Date object as a human-readable string (e.g., "January 15, 2025").
 * @param {Date} date - The date to format
 * @returns {string} Formatted date string
 */
function formatDate(date) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
    return '';
  }

  const options = { year: 'numeric', month: 'long', day: 'numeric' };
  return date.toLocaleDateString('en-US', options);
}

/**
 * Formats a Date object as ISO date string (YYYY-MM-DD).
 * @param {Date} date - The date to format
 * @returns {string} ISO date string
 */
function formatDateISO(date) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
    return '';
  }

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Parses reminder offsets from a string or array.
 * Accepts comma-separated values like "3,7,14" or "3 days before,1 week before".
 * @param {string|Array} offsets - The reminder offsets to parse
 * @returns {number[]} Array of integer day offsets
 */
function parseReminderOffsets(offsets) {
  if (!offsets) {
    return [];
  }

  if (Array.isArray(offsets)) {
    return offsets.map(o => parseInt(o, 10)).filter(n => !isNaN(n));
  }

  const parts = String(offsets).split(',').map(s => s.trim());
  const result = [];

  for (const part of parts) {
    // Try direct integer parse first
    const num = parseInt(part, 10);
    if (!isNaN(num)) {
      result.push(num);
      continue;
    }

    // Try extracting number from human-readable format like "3 days before"
    const match = part.match(/(\d+)\s*(day|week)/i);
    if (match) {
      const value = parseInt(match[1], 10);
      const unit = match[2].toLowerCase();
      result.push(unit === 'week' ? value * 7 : value);
    }
  }

  return result;
}

/**
 * Substitutes tokens in a template string with values from a data object.
 * Tokens are in the format {{TOKEN_NAME}}.
 * @param {string} template - The template string with tokens
 * @param {Object} values - Key-value pairs for substitution
 * @returns {string} The template with tokens replaced
 */
function substituteTokens(template, values) {
  if (!template) {
    return '';
  }

  let result = template;

  for (const [key, value] of Object.entries(values)) {
    const token = `{{${key}}}`;
    result = result.split(token).join(value || '');
  }

  return result;
}

/**
 * Calculates the number of days between two dates.
 * @param {Date} fromDate - The start date
 * @param {Date} toDate - The end date
 * @returns {number} Number of days (positive if toDate is in the future)
 */
function daysBetween(fromDate, toDate) {
  if (!fromDate || !toDate) {
    return 0;
  }

  const msPerDay = 24 * 60 * 60 * 1000;
  const from = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
  const to = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());

  return Math.round((to - from) / msPerDay);
}

/**
 * Checks if two dates are the same calendar day.
 * @param {Date} date1 - First date
 * @param {Date} date2 - Second date
 * @returns {boolean} True if same day
 */
function isSameDay(date1, date2) {
  if (!date1 || !date2) {
    return false;
  }

  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}

/**
 * Generates a Google Drive folder URL from a folder ID.
 * @param {string} folderId - The Google Drive folder ID
 * @returns {string} The folder URL
 */
function folderIdToUrl(folderId) {
  if (!folderId) {
    return '';
  }
  return `https://drive.google.com/drive/folders/${folderId}`;
}

/**
 * Parses a comma-separated string into an array of trimmed values.
 * @param {string} str - The comma-separated string
 * @returns {string[]} Array of trimmed values
 */
function parseCommaSeparated(str) {
  if (!str) {
    return [];
  }

  return String(str)
    .split(',')
    .map(s => s.trim())
    .filter(s => s !== '');
}

/**
 * Joins an array into a comma-separated string.
 * @param {string[]} arr - The array to join
 * @returns {string} Comma-separated string
 */
function joinCommaSeparated(arr) {
  if (!arr || !Array.isArray(arr)) {
    return '';
  }

  return arr.filter(s => s).join(', ');
}

/**
 * Gets the start of today (midnight) in the script's timezone.
 * @returns {Date} Today at midnight
 */
function getStartOfToday() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

/**
 * Safely parses a value as a Date object.
 * @param {*} value - The value to parse
 * @returns {Date|null} Parsed Date or null if invalid
 */
function parseDate(value) {
  if (!value) {
    return null;
  }

  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }

  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Infers the school year from a deadline date.
 * School year format is "YY_YY" (e.g., "25_26" for 2025-2026).
 * @param {Date} deadline - The project deadline date
 * @param {number} startMonth - Month when school year starts (1-12, default: 7 for July)
 * @returns {string} School year in YY_YY format
 */
function inferSchoolYear(deadline, startMonth = 7) {
  if (!deadline || !(deadline instanceof Date) || isNaN(deadline.getTime())) {
    return '';
  }

  const year = deadline.getFullYear();
  const month = deadline.getMonth() + 1; // Convert to 1-based

  // If deadline is in/after the start month, it's the "new" school year
  // e.g., July 2026+ → '26_27, before July 2026 → '25_26
  if (month >= startMonth) {
    const startYear = year % 100;
    const endYear = (year + 1) % 100;
    return `${padNumber(startYear, 2)}_${padNumber(endYear, 2)}`;
  } else {
    const startYear = (year - 1) % 100;
    const endYear = year % 100;
    return `${padNumber(startYear, 2)}_${padNumber(endYear, 2)}`;
  }
}

/**
 * Creates a display title for a project in the format "[ID] Name".
 * @param {string} projectId - The project ID
 * @param {string} projectName - The project name
 * @returns {string} Display title
 */
function formatProjectTitle(projectId, projectName) {
  if (projectId && projectName) {
    return `${projectName} [${projectId}]`;
  }
  return projectName || projectId || '';
}

/**
 * Pads a number with leading zeros.
 * @param {number} num - The number to pad
 * @param {number} length - Desired string length
 * @returns {string} Padded number string
 */
function padNumber(num, length) {
  return String(num).padStart(length, '0');
}

/**
 * Validates that an email address has a basic valid format.
 * @param {string} email - The email to validate
 * @returns {boolean} True if valid format
 */
function isValidEmail(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }
  // Basic email regex - not exhaustive but catches obvious issues
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim());
}

/**
 * Extracts email addresses from a string that may contain names and emails.
 * Handles formats like "John Doe <john@example.com>" or just "john@example.com".
 * @param {string} str - The string containing emails
 * @returns {string[]} Array of email addresses
 */
function extractEmails(str) {
  if (!str) {
    return [];
  }

  const emails = [];
  const parts = parseCommaSeparated(str);

  for (const part of parts) {
    // Check for "Name <email>" format
    const angleMatch = part.match(/<([^>]+)>/);
    if (angleMatch) {
      const email = angleMatch[1].trim();
      if (isValidEmail(email)) {
        emails.push(email);
      }
      continue;
    }

    // Check if the part itself is an email
    if (isValidEmail(part)) {
      emails.push(part.trim());
    }
  }

  return emails;
}

