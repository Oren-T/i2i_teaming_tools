/**
 * Codes class wrapping the Codes sheet.
 * Provides dropdown values for categories, statuses, and reminder offsets.
 */
class Codes {
  /**
   * Creates a new Codes instance.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Codes sheet
   */
  constructor(sheet) {
    this.sheet = sheet;
    this.data = null;
    this.headerMap = null;
    this.loadData();
  }

  /**
   * Loads data from the Codes sheet into memory.
   */
  loadData() {
    if (!this.sheet) {
      throw new Error('Codes sheet not found');
    }

    const dataRange = this.sheet.getDataRange();
    this.data = dataRange.getValues();

    if (this.data.length === 0) {
      DEBUG && console.log('Codes: Sheet is empty');
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

    DEBUG && console.log(`Codes: Loaded ${this.data.length - 1} rows with ${this.headerMap.size} columns`);
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
   * Gets all non-empty values from a column by header name.
   * @param {string} headerName - The column header name
   * @returns {string[]} Array of values from that column
   */
  getColumnValues(headerName) {
    const colIndex = this.getColumnIndex(headerName);
    if (colIndex === undefined) {
      DEBUG && console.log(`Codes: Column "${headerName}" not found`);
      return [];
    }

    const values = [];
    for (let i = 1; i < this.data.length; i++) {
      const value = this.data[i][colIndex];
      if (value !== null && value !== undefined && String(value).trim() !== '') {
        values.push(String(value).trim());
      }
    }

    return values;
  }

  /**
   * Gets all project categories.
   * @returns {string[]} Array of category values (e.g., ["LCAP", "SPSA", "Community School", ...])
   */
  getCategories() {
    return this.getColumnValues(CODES_COLUMNS.CATEGORY);
  }

  /**
   * Gets all project statuses.
   * @returns {string[]} Array of status values (e.g., ["Not started", "On track", "Completed", ...])
   */
  getStatuses() {
    return this.getColumnValues(CODES_COLUMNS.STATUS);
  }

  /**
   * Gets reminder day offsets (integer values).
   * @returns {number[]} Array of integer day offsets (e.g., [3, 7, 14])
   */
  getReminderOffsets() {
    const values = this.getColumnValues(CODES_COLUMNS.REMINDER_DAYS_OFFSET);
    return values.map(v => parseInt(v, 10)).filter(n => !isNaN(n));
  }

  /**
   * Gets human-readable reminder labels.
   * @returns {string[]} Array of labels (e.g., ["3 days before", "1 week before", ...])
   */
  getReminderLabels() {
    return this.getColumnValues(CODES_COLUMNS.REMINDER_DAYS_READABLE);
  }

  /**
   * Gets reminder offset/label pairs.
   * @returns {Object[]} Array of {offset, label} objects
   */
  getReminderOptions() {
    const offsets = this.getReminderOffsets();
    const labels = this.getReminderLabels();
    const options = [];

    const maxLen = Math.max(offsets.length, labels.length);
    for (let i = 0; i < maxLen; i++) {
      if (offsets[i] !== undefined) {
        options.push({
          offset: offsets[i],
          label: labels[i] || `${offsets[i]} days before`
        });
      }
    }

    return options;
  }

  /**
   * Maps a human-readable reminder label to its integer offset.
   * @param {string} label - The readable label
   * @returns {number|null} The integer offset or null if not found
   */
  labelToOffset(label) {
    if (!label) return null;

    const labels = this.getReminderLabels();
    const offsets = this.getReminderOffsets();

    const index = labels.findIndex(l =>
      String(l).trim().toLowerCase() === String(label).trim().toLowerCase()
    );

    if (index !== -1 && offsets[index] !== undefined) {
      return offsets[index];
    }

    // Try parsing directly if it's just a number
    const num = parseInt(label, 10);
    return isNaN(num) ? null : num;
  }

  /**
   * Maps an integer offset to its human-readable label.
   * @param {number} offset - The integer offset
   * @returns {string} The readable label or a default format
   */
  offsetToLabel(offset) {
    if (offset === null || offset === undefined) return '';

    const offsets = this.getReminderOffsets();
    const labels = this.getReminderLabels();

    const index = offsets.indexOf(offset);
    if (index !== -1 && labels[index]) {
      return labels[index];
    }

    // Generate default label
    if (offset === 7) return '1 week before';
    if (offset === 14) return '2 weeks before';
    if (offset === 21) return '3 weeks before';
    return `${offset} days before`;
  }

  /**
   * Checks if a category value is valid.
   * @param {string} category - The category to check
   * @returns {boolean} True if valid
   */
  isValidCategory(category) {
    if (!category) return false;
    const categories = this.getCategories();
    return categories.some(c =>
      String(c).trim().toLowerCase() === String(category).trim().toLowerCase()
    );
  }

  /**
   * Checks if a status value is valid.
   * @param {string} status - The status to check
   * @returns {boolean} True if valid
   */
  isValidStatus(status) {
    if (!status) return false;
    const statuses = this.getStatuses();
    return statuses.some(s =>
      String(s).trim().toLowerCase() === String(status).trim().toLowerCase()
    );
  }

  /**
   * Gets the default category value.
   * @returns {string} Default category (first in list or DEFAULTS.CATEGORY)
   */
  getDefaultCategory() {
    const categories = this.getCategories();
    // Check if LCAP is in the list
    const lcap = categories.find(c => c.toUpperCase() === 'LCAP');
    return lcap || categories[0] || DEFAULTS.CATEGORY;
  }

  /**
   * Gets the default reminder offsets.
   * @returns {number[]} Default offset values
   */
  getDefaultReminderOffsets() {
    const offsets = this.getReminderOffsets();
    return offsets.length > 0 ? offsets : DEFAULTS.REMINDER_OFFSETS;
  }
}

