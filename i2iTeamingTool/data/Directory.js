/**
 * Directory class wrapping the Staff Directory sheet.
 * Provides staff name/email lookups and active staff lists.
 */
class Directory {
  /**
   * Creates a new Directory instance.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Directory sheet
   */
  constructor(sheet) {
    this.sheet = sheet;
    this.data = null;
    this.headerMap = null;
    this.nameToEmail = new Map();
    this.emailToName = new Map();
    this.loadData();
  }

  /**
   * Loads data from the Directory sheet into memory.
   */
  loadData() {
    if (!this.sheet) {
      throw new Error('Directory sheet not found');
    }

    const dataRange = this.sheet.getDataRange();
    this.data = dataRange.getValues();

    if (this.data.length === 0) {
      DEBUG && console.log('Directory: Sheet is empty');
      return;
    }

    // Build header map from Row 1
    this.headerMap = new Map();
    const headers = this.data[0];
    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i]).trim();
      if (header) {
        this.headerMap.set(header, i);
      }
    }

    // Build lookup maps
    const nameCol = this.headerMap.get(DIRECTORY_COLUMNS.NAME);
    const emailCol = this.headerMap.get(DIRECTORY_COLUMNS.EMAIL);

    if (nameCol === undefined || emailCol === undefined) {
      console.warn('Directory: Missing required columns (Name, Email Address)');
      return;
    }

    // Skip header row, process data rows
    for (let i = 1; i < this.data.length; i++) {
      const name = String(this.data[i][nameCol] || '').trim();
      const email = String(this.data[i][emailCol] || '').trim().toLowerCase();

      if (name && email) {
        this.nameToEmail.set(name.toLowerCase(), email);
        this.emailToName.set(email, name);
      }
    }

    DEBUG && console.log(`Directory: Loaded ${this.nameToEmail.size} staff entries`);
  }

  /**
   * Gets the column index for a header name.
   * @param {string} header - The header name
   * @returns {number|undefined} Column index (0-based) or undefined
   */
  getColumnIndex(header) {
    return this.headerMap ? this.headerMap.get(header) : undefined;
  }

  /**
   * Looks up a staff member's email by their name.
   * @param {string} name - The staff member's name
   * @returns {string|null} Email address or null if not found
   */
  getEmailByName(name) {
    if (!name) return null;
    const email = this.nameToEmail.get(String(name).trim().toLowerCase());
    return email || null;
  }

  /**
   * Looks up a staff member's name by their email.
   * @param {string} email - The staff member's email
   * @returns {string|null} Name or null if not found
   */
  getNameByEmail(email) {
    if (!email) return null;
    const name = this.emailToName.get(String(email).trim().toLowerCase());
    return name || null;
  }

  /**
   * Gets all staff names from the directory.
   * @returns {string[]} Array of staff names
   */
  getAllStaffNames() {
    const names = [];
    const nameCol = this.getColumnIndex(DIRECTORY_COLUMNS.NAME);

    if (nameCol === undefined) return names;

    for (let i = 1; i < this.data.length; i++) {
      const name = String(this.data[i][nameCol] || '').trim();
      if (name) {
        names.push(name);
      }
    }

    return names;
  }

  /**
   * Gets all staff emails from the directory.
   * @returns {string[]} Array of email addresses
   */
  getAllStaffEmails() {
    const emails = [];
    const emailCol = this.getColumnIndex(DIRECTORY_COLUMNS.EMAIL);

    if (emailCol === undefined) return emails;

    for (let i = 1; i < this.data.length; i++) {
      const email = String(this.data[i][emailCol] || '').trim();
      if (email) {
        emails.push(email);
      }
    }

    return emails;
  }

  /**
   * Gets names of active staff members.
   * If a Permissions column exists, filters to those with non-empty permissions.
   * @returns {string[]} Array of active staff names
   */
  getActiveStaffNames() {
    const names = [];
    const nameCol = this.getColumnIndex(DIRECTORY_COLUMNS.NAME);
    const permCol = this.getColumnIndex(DIRECTORY_COLUMNS.PERMISSIONS);

    if (nameCol === undefined) return names;

    for (let i = 1; i < this.data.length; i++) {
      const name = String(this.data[i][nameCol] || '').trim();

      // If there's a permissions column, check if staff is active
      if (permCol !== undefined) {
        const perm = String(this.data[i][permCol] || '').trim().toLowerCase();
        // Skip if permissions indicate inactive (empty or "none" or "no access")
        if (!perm || perm === 'none' || perm === 'no access') {
          continue;
        }
      }

      if (name) {
        names.push(name);
      }
    }

    return names;
  }

  /**
   * Gets emails of active staff members.
   * If a Permissions column exists, filters to those with non-empty permissions.
   * @returns {string[]} Array of active staff email addresses
   */
  getActiveStaffEmails() {
    const emails = [];
    const emailCol = this.getColumnIndex(DIRECTORY_COLUMNS.EMAIL);
    const permCol = this.getColumnIndex(DIRECTORY_COLUMNS.PERMISSIONS);

    if (emailCol === undefined) return emails;

    for (let i = 1; i < this.data.length; i++) {
      const email = String(this.data[i][emailCol] || '').trim();

      // If there's a permissions column, check if staff is active
      if (permCol !== undefined) {
        const perm = String(this.data[i][permCol] || '').trim().toLowerCase();
        if (!perm || perm === 'none' || perm === 'no access') {
          continue;
        }
      }

      if (email) {
        emails.push(email);
      }
    }

    return emails;
  }

  /**
   * Checks if a name exists in the directory.
   * @param {string} name - The name to check
   * @returns {boolean} True if name exists
   */
  hasName(name) {
    if (!name) return false;
    return this.nameToEmail.has(String(name).trim().toLowerCase());
  }

  /**
   * Checks if an email exists in the directory.
   * @param {string} email - The email to check
   * @returns {boolean} True if email exists
   */
  hasEmail(email) {
    if (!email) return false;
    return this.emailToName.has(String(email).trim().toLowerCase());
  }

  /**
   * Resolves a name or email to an email address.
   * If input is already a valid email, returns it.
   * If input is a name, looks up the email.
   * @param {string} nameOrEmail - Name or email to resolve
   * @returns {string|null} Email address or null if not found
   */
  resolveToEmail(nameOrEmail) {
    if (!nameOrEmail) return null;

    const trimmed = String(nameOrEmail).trim();

    // Check if it's already an email
    if (isValidEmail(trimmed)) {
      return trimmed.toLowerCase();
    }

    // Try to look up by name
    return this.getEmailByName(trimmed);
  }

  /**
   * Resolves multiple names/emails to email addresses.
   * Accepts comma-separated string or array.
   * @param {string|string[]} namesOrEmails - Names or emails to resolve
   * @returns {string[]} Array of resolved email addresses
   */
  resolveAllToEmails(namesOrEmails) {
    const inputs = Array.isArray(namesOrEmails)
      ? namesOrEmails
      : parseCommaSeparated(namesOrEmails);

    const emails = [];
    for (const input of inputs) {
      const email = this.resolveToEmail(input);
      if (email) {
        emails.push(email);
      }
    }

    return emails;
  }
}

