/**
 * Config class wrapping the Config sheet with typed getters.
 * Provides access to district-specific configuration values.
 */
class Config {
  /**
   * Creates a new Config instance.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Config sheet
   */
  constructor(sheet) {
    this.sheet = sheet;
    this.data = null;
    this.keyValueMap = null;
    this.loadData();
  }

  /**
   * Loads data from the Config sheet into memory.
   */
  loadData() {
    if (!this.sheet) {
      throw new Error('Config sheet not found');
    }

    const dataRange = this.sheet.getDataRange();
    this.data = dataRange.getValues();

    // Build key-value map (Column A = key, Column B = value)
    this.keyValueMap = new Map();
    for (let i = 0; i < this.data.length; i++) {
      const key = this.data[i][0];
      const value = this.data[i][1];
      if (key) {
        this.keyValueMap.set(String(key).trim(), value);
      }
    }

    DEBUG && console.log(`Config: Loaded ${this.keyValueMap.size} configuration entries`);
  }

  /**
   * Gets a raw value from the config by key.
   * @param {string} key - The config key
   * @returns {*} The value, or undefined if not found
   */
  get(key) {
    return this.keyValueMap.get(key);
  }

  /**
   * Gets the raw data array (for validation).
   * @returns {Array[]} The raw config data
   */
  getRawData() {
    return this.data || [];
  }

  // ===== TYPED GETTERS =====

  /**
   * Gets the District ID (e.g., "NUSD").
   * @returns {string} District ID
   */
  get districtId() {
    return String(this.get('District ID') || '').trim();
  }

  /**
   * Gets the next serial number for project IDs.
   * @returns {number} Next serial number
   */
  get nextSerial() {
    const value = this.get('Next Serial');
    return parseInt(value, 10) || 1;
  }

  /**
   * Gets the Parent Folder ID where project folders are created.
   * @returns {string} Google Drive folder ID
   */
  get parentFolderId() {
    return String(this.get('Parent Folder ID') || '').trim();
  }

  /**
   * Gets the Root Folder ID for this Project Management Tool instance.
   * This is the top-level folder that contains Templates, Backups,
   * the Project Folders subfolder, and the main spreadsheet.
   * @returns {string} Google Drive folder ID
   */
  get rootFolderId() {
    return String(this.get('Root Folder ID') || '').trim();
  }

  /**
   * Gets the Backups Folder ID where weekly spreadsheet backups are stored.
   * @returns {string} Google Drive folder ID
   */
  get backupsFolderId() {
    return String(this.get('Backups Folder ID') || '').trim();
  }

  /**
   * Gets the Project Template ID (Google Sheets file to copy).
   * @returns {string} Google Sheets file ID
   */
  get projectTemplateId() {
    return String(this.get('Project Template ID') || '').trim();
  }

  /**
   * Gets the Google Form ID for project submissions.
   * @returns {string} Google Form ID
   */
  get formId() {
    return String(this.get('Form ID') || '').trim();
  }

  /**
   * Gets the error notification email addresses.
   * @returns {string[]} Array of email addresses
   */
  get errorEmailAddresses() {
    const value = this.get('Error Email Addresses') || '';
    return parseCommaSeparated(value);
  }

  /**
   * Gets the New Project email template Doc ID.
   * @returns {string} Google Doc ID
   */
  get emailTemplateNewProject() {
    return String(this.get('Email Template - New Project') || '').trim();
  }

  /**
   * Gets the Reminder email template Doc ID.
   * @returns {string} Google Doc ID
   */
  get emailTemplateReminder() {
    return String(this.get('Email Template - Reminder') || '').trim();
  }

  /**
   * Gets the Status Change email template Doc ID.
   * @returns {string} Google Doc ID
   */
  get emailTemplateStatusChange() {
    return String(this.get('Email Template - Status Change') || '').trim();
  }

  /**
   * Gets the Project Update email template Doc ID.
   * @returns {string} Google Doc ID
   */
  get emailTemplateUpdate() {
    return String(this.get('Email Template - Project Update') || '').trim();
  }

  /**
   * Gets the Project Cancellation email template Doc ID.
   * @returns {string} Google Doc ID
   */
  get emailTemplateCancellation() {
    return String(this.get('Email Template - Project Cancellation') || '').trim();
  }

  /**
   * Gets the month when the school year starts (1-12).
   * Defaults to 7 (July) if not configured.
   * @returns {number} Start month (1 = January, 7 = July, etc.)
   */
  get schoolYearStartMonth() {
    const value = this.get('School Year Start Month');
    const parsed = parseInt(value, 10);
    // Validate range 1-12, default to 7 (July)
    if (isNaN(parsed) || parsed < 1 || parsed > 12) {
      return 7;
    }
    return parsed;
  }

  // ===== SERIAL NUMBER MANAGEMENT =====

  /**
   * Gets the next serial number and increments it in the sheet.
   * Should be called within a script lock to prevent race conditions.
   * @returns {number} The current serial number (before increment)
   */
  getAndIncrementSerial() {
    const currentSerial = this.nextSerial;
    const newSerial = currentSerial + 1;

    // Find the row with "Next Serial" and update Column B
    for (let i = 0; i < this.data.length; i++) {
      if (String(this.data[i][0]).trim() === 'Next Serial') {
        const rowNumber = i + 1; // 1-based row number
        this.sheet.getRange(rowNumber, 2).setValue(newSerial);
        SpreadsheetApp.flush(); // Force write inside the lock to prevent race conditions

        // Update local cache
        this.data[i][1] = newSerial;
        this.keyValueMap.set('Next Serial', newSerial);

        DEBUG && console.log(`Config: Incremented serial from ${currentSerial} to ${newSerial}`);
        break;
      }
    }

    return currentSerial;
  }

  /**
   * Sets a config value by key.
   * @param {string} key - The config key
   * @param {*} value - The value to set
   */
  set(key, value) {
    for (let i = 0; i < this.data.length; i++) {
      if (String(this.data[i][0]).trim() === key) {
        const rowNumber = i + 1;
        this.sheet.getRange(rowNumber, 2).setValue(value);

        // Update local cache
        this.data[i][1] = value;
        this.keyValueMap.set(key, value);

        DEBUG && console.log(`Config: Set "${key}" to "${value}"`);
        return;
      }
    }

    // Key not found - could add new row, but for now just log
    console.warn(`Config: Key "${key}" not found in Config sheet`);
  }
}

