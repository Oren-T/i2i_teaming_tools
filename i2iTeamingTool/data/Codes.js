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
    this.headerRow = CODES_LAYOUT.HEADER_ROW;
    this.categories = [];
    this.statuses = [];
    this.reminderOffsets = [];
    this.reminderLabels = [];
    this.loadData();
  }

  /**
   * Loads data from the Codes sheet into memory.
   */
  loadData() {
    if (!this.sheet) {
      throw new Error('Codes sheet not found');
    }

    this.lastRow = this.sheet.getLastRow();

    if (!this.headerRow || this.lastRow < this.headerRow) {
      DEBUG && console.log('Codes: Sheet is missing header row or data');
      this.categories = [];
      this.statuses = [];
      this.reminderOffsets = [];
      this.reminderLabels = [];
      return;
    }

    this.categories = this.readColumnValues(CODES_LAYOUT.CATEGORY_COL, CODES_COLUMNS.CATEGORY);
    this.statuses = this.readColumnValues(CODES_LAYOUT.STATUS_COL, CODES_COLUMNS.STATUS);
    const reminders = this.readReminderColumns();
    this.reminderOffsets = reminders.offsets;
    this.reminderLabels = reminders.labels;

    DEBUG && console.log(
      `Codes: Loaded ${this.categories.length} categories, ` +
      `${this.statuses.length} statuses, ${this.reminderOffsets.length} reminder offsets`
    );
  }

  /**
   * Reads a single anchored column, validates its header, and filters values.
   * @param {number} columnIndex - 1-based column index
   * @param {string} expectedHeader - Expected header text
   * @returns {string[]} Filtered values
   */
  readColumnValues(columnIndex, expectedHeader) {
    const numRows = this.lastRow - this.headerRow + 1;
    if (numRows <= 0) {
      return [];
    }

    const rangeValues = this.sheet.getRange(this.headerRow, columnIndex, numRows, 1).getValues();
    this.ensureHeaderMatches(rangeValues[0][0], expectedHeader, columnIndex);

    const results = [];
    const seen = new Set();

    for (let i = 1; i < rangeValues.length; i++) {
      const normalized = this.normalizeValue(rangeValues[i][0]);
      if (!normalized) {
        continue;
      }

      const dedupeKey = normalized.toLowerCase();
      if (seen.has(dedupeKey)) {
        continue;
      }

      seen.add(dedupeKey);
      results.push(normalized);
    }

    return results;
  }

  /**
   * Reads the reminder offset + label columns together to keep them aligned.
   * @returns {{offsets: number[], labels: string[]}}
   */
  readReminderColumns() {
    const numRows = this.lastRow - this.headerRow + 1;
    if (numRows <= 0) {
      return { offsets: [], labels: [] };
    }

    const rangeValues = this.sheet.getRange(
      this.headerRow,
      CODES_LAYOUT.REMINDER_OFFSET_COL,
      numRows,
      2
    ).getValues();

    this.ensureHeaderMatches(
      rangeValues[0][0],
      CODES_COLUMNS.REMINDER_DAYS_OFFSET,
      CODES_LAYOUT.REMINDER_OFFSET_COL
    );
    this.ensureHeaderMatches(
      rangeValues[0][1],
      CODES_COLUMNS.REMINDER_DAYS_READABLE,
      CODES_LAYOUT.REMINDER_LABEL_COL
    );

    const offsets = [];
    const labels = [];
    const seen = new Set();

    for (let i = 1; i < rangeValues.length; i++) {
      const [rawOffset, rawLabel] = rangeValues[i];
      const offset = parseInt(rawOffset, 10);

      if (isNaN(offset)) {
        continue;
      }

      if (seen.has(offset)) {
        continue;
      }

      seen.add(offset);
      offsets.push(offset);
      labels.push(this.normalizeValue(rawLabel));
    }

    return { offsets, labels };
  }

  /**
   * Ensures the header cell matches the expected text.
   * @param {any} actualValue - Actual header cell value
   * @param {string} expected - Expected header text
   * @param {number} columnIndex - Column index (1-based)
   */
  ensureHeaderMatches(actualValue, expected, columnIndex) {
    const normalized = this.normalizeValue(actualValue);
    if (normalized === expected) {
      return;
    }

    const columnLetter = this.columnToLetter(columnIndex);
    throw new Error(`Codes: Expected "${expected}" in cell ${columnLetter}${this.headerRow} but found "${normalized || ''}"`);
  }

  /**
   * Normalizes a value by converting to string and trimming.
   * @param {any} value - Value to normalize
   * @returns {string} Normalized string value
   */
  normalizeValue(value) {
    if (value === null || value === undefined) {
      return '';
    }
    return String(value).trim();
  }

  /**
   * Converts a column number to its spreadsheet letter (1 -> A).
   * @param {number} columnIndex - Column index
   * @returns {string} Column letter
   */
  columnToLetter(columnIndex) {
    let index = columnIndex;
    let letter = '';

    while (index > 0) {
      const remainder = (index - 1) % 26;
      letter = String.fromCharCode(65 + remainder) + letter;
      index = Math.floor((index - 1) / 26);
    }

    return letter;
  }

  /**
   * Gets all project categories.
   * @returns {string[]} Array of category values (e.g., ["LCAP", "SPSA", "Community School", ...])
   */
  getCategories() {
    return [...this.categories];
  }

  /**
   * Gets all project statuses.
   * @returns {string[]} Array of status values (e.g., ["Not started", "On track", "Complete", ...])
   */
  getStatuses() {
    return [...this.statuses];
  }

  /**
   * Gets reminder day offsets (integer values).
   * @returns {number[]} Array of integer day offsets (e.g., [3, 7, 14])
   */
  getReminderOffsets() {
    return [...this.reminderOffsets];
  }

  /**
   * Gets human-readable reminder labels.
   * @returns {string[]} Array of labels (e.g., ["3 days before", "1 week before", ...])
   */
  getReminderLabels() {
    return [...this.reminderLabels];
  }

  /**
   * Gets reminder offset/label pairs.
   * @returns {Object[]} Array of {offset, label} objects
   */
  getReminderOptions() {
    const options = [];

    for (let i = 0; i < this.reminderOffsets.length; i++) {
      const offset = this.reminderOffsets[i];
      const label = this.reminderLabels[i] || this.buildDefaultReminderLabel(offset);
      options.push({ offset, label });
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

    const normalized = this.normalizeValue(label).toLowerCase();

    for (let i = 0; i < this.reminderLabels.length; i++) {
      const currentLabel = this.reminderLabels[i];
      if (currentLabel && currentLabel.toLowerCase() === normalized) {
        return this.reminderOffsets[i];
      }
    }

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

    const index = this.reminderOffsets.indexOf(offset);
    if (index !== -1) {
      const label = this.reminderLabels[index];
      if (label) {
        return label;
      }
    }

    return this.buildDefaultReminderLabel(offset);
  }

  /**
   * Builds a default reminder label for a given offset.
   * @param {number} offset - Day offset
   * @returns {string} Default label text
   */
  buildDefaultReminderLabel(offset) {
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
    const normalized = this.normalizeValue(category).toLowerCase();
    return this.categories.some(c => c.toLowerCase() === normalized);
  }

  /**
   * Checks if a status value is valid.
   * @param {string} status - The status to check
   * @returns {boolean} True if valid
   */
  isValidStatus(status) {
    if (!status) return false;
    const normalized = this.normalizeValue(status).toLowerCase();
    return this.statuses.some(s => s.toLowerCase() === normalized);
  }

  /**
   * Gets the default category value.
   * @returns {string} Default category (first in list or DEFAULTS.CATEGORY)
   */
  getDefaultCategory() {
    const lcap = this.categories.find(c => c.toUpperCase() === 'LCAP');
    return lcap || this.categories[0] || DEFAULTS.CATEGORY;
  }

  /**
   * Gets the default reminder offsets.
   * @returns {number[]} Default offset values
   */
  getDefaultReminderOffsets() {
    return this.reminderOffsets.length > 0 ? [...this.reminderOffsets] : [...DEFAULTS.REMINDER_OFFSETS];
  }

  /**
   * Gets the default reminder labels (human-readable format).
   * Falls back to auto-generated labels if none are defined in the Codes sheet.
   * @returns {string[]} Default label values (e.g., ["3 days before", "1 week before"])
   */
  getDefaultReminderLabels() {
    if (this.reminderLabels.length > 0) {
      return [...this.reminderLabels];
    }

    // Fall back to generating labels from default offsets
    return DEFAULTS.REMINDER_OFFSETS.map(offset => this.buildDefaultReminderLabel(offset));
  }
}

