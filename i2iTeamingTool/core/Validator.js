/**
 * Validator class for validating environment at startup.
 * Checks config keys, project columns, and file access.
 */
class Validator {
  /**
   * Creates a new Validator instance.
   * @param {Config} config - The Config instance
   * @param {ProjectSheet} projectSheet - The ProjectSheet instance
   */
  constructor(config, projectSheet) {
    this.config = config;
    this.projectSheet = projectSheet;
    this.errors = [];
  }

  /**
   * Runs validations and throws if any errors found.
   * @param {Object} [options] - Validation options
   * @param {boolean} [options.includeFileAccess=false] - Whether to include external file checks
   * @throws {Error} Aggregated error message if validation fails
   */
  validate(options = {}) {
    const { includeFileAccess = false } = options;
    this.errors = [];

    this.validateConfigKeys();
    this.validateProjectColumns();

    if (includeFileAccess) {
      this.validateFileAccess();
    }

    if (this.errors.length > 0) {
      const errorMessage = `Validation failed with ${this.errors.length} error(s):\n` +
        this.errors.map((e, i) => `  ${i + 1}. ${e}`).join('\n');
      console.error(errorMessage);
      throw new Error(errorMessage);
    }

    DEBUG && console.log('Validator: All validations passed');
  }

  /**
   * Validates that all required config keys exist in the Config sheet.
   */
  validateConfigKeys() {
    DEBUG && console.log('Validator: Checking config keys...');

    const configData = this.config.getRawData();
    const existingKeys = new Set();

    // Build set of existing keys (Column A values)
    for (const row of configData) {
      if (row[0]) {
        existingKeys.add(String(row[0]).trim());
      }
    }

    // Check for missing required keys
    for (const key of REQUIRED_CONFIG_KEYS) {
      if (!existingKeys.has(key)) {
        this.errors.push(`Missing required config key: "${key}"`);
      }
    }

    DEBUG && console.log(`Validator: Found ${existingKeys.size} config keys, ${REQUIRED_CONFIG_KEYS.length} required`);
  }

  /**
   * Validates that all required project columns exist in Row 2 and are unique.
   */
  validateProjectColumns() {
    DEBUG && console.log('Validator: Checking project columns...');

    const columnKeys = this.projectSheet.getColumnKeys();

    if (!columnKeys || columnKeys.length === 0) {
      this.errors.push('Project sheet Row 2 (internal keys) is empty');
      return;
    }

    // Check for duplicates
    const seen = new Set();
    const duplicates = new Set();

    for (const key of columnKeys) {
      if (!key) continue;
      const trimmed = String(key).trim();
      if (seen.has(trimmed)) {
        duplicates.add(trimmed);
      }
      seen.add(trimmed);
    }

    if (duplicates.size > 0) {
      this.errors.push(`Duplicate column keys in Row 2: ${Array.from(duplicates).join(', ')}`);
    }

    // Check for missing required columns
    for (const required of REQUIRED_PROJECT_COLUMNS) {
      if (!seen.has(required)) {
        this.errors.push(`Missing required project column: "${required}"`);
      }
    }

    DEBUG && console.log(`Validator: Found ${seen.size} column keys, ${REQUIRED_PROJECT_COLUMNS.length} required`);
  }

  /**
   * Validates that required external files are accessible.
   */
  validateFileAccess() {
    DEBUG && console.log('Validator: Checking file access...');

    // Check Root Folder
    const rootFolderId = this.config.rootFolderId;
    if (rootFolderId) {
      try {
        withBackoff(() => DriveApp.getFolderById(rootFolderId));
        DEBUG && console.log('Validator: Root folder accessible');
      } catch (e) {
        this.errors.push(`Cannot access Root Folder (ID: ${rootFolderId}): ${e.message}`);
      }
    } else {
      this.errors.push('Root Folder ID is not configured');
    }

    // Check Parent Folder (Project Folders)
    const parentFolderId = this.config.parentFolderId;
    if (parentFolderId) {
      try {
        withBackoff(() => DriveApp.getFolderById(parentFolderId));
        DEBUG && console.log('Validator: Parent folder accessible');
      } catch (e) {
        this.errors.push(`Cannot access Parent Folder (ID: ${parentFolderId}): ${e.message}`);
      }
    } else {
      this.errors.push('Parent Folder ID is not configured');
    }

    // Check Project Template
    const templateId = this.config.projectTemplateId;
    if (templateId) {
      try {
        withBackoff(() => DriveApp.getFileById(templateId));
        DEBUG && console.log('Validator: Project template accessible');
      } catch (e) {
        this.errors.push(`Cannot access Project Template (ID: ${templateId}): ${e.message}`);
      }
    } else {
      this.errors.push('Project Template ID is not configured');
    }

    // Check Form (optional - may not be configured initially)
    const formId = this.config.formId;
    if (formId) {
      try {
        withBackoff(() => FormApp.openById(formId));
        DEBUG && console.log('Validator: Form accessible');
      } catch (e) {
        this.errors.push(`Cannot access Form (ID: ${formId}): ${e.message}`);
      }
    } else {
      DEBUG && console.log('Validator: Form ID not configured (optional)');
    }

    // Check email templates (required - notifications won't work without them)
    const emailTemplateKeys = [
      { key: 'emailTemplateNewProject', name: 'New Project' },
      { key: 'emailTemplateReminder', name: 'Reminder' },
      { key: 'emailTemplateStatusChange', name: 'Status Change' },
      { key: 'emailTemplateUpdate', name: 'Project Update' },
      { key: 'emailTemplateCancellation', name: 'Project Cancellation' }
    ];

    for (const { key, name } of emailTemplateKeys) {
      const docId = this.config[key];
      if (docId) {
        try {
          withBackoff(() => DocumentApp.openById(docId));
          DEBUG && console.log(`Validator: Email template "${name}" accessible`);
        } catch (e) {
          this.errors.push(`Cannot access Email Template - ${name} (ID: ${docId}): ${e.message}`);
        }
      } else {
        this.errors.push(`Email Template - ${name} is not configured`);
      }
    }
  }

  /**
   * Returns the list of validation errors.
   * @returns {string[]} Array of error messages
   */
  getErrors() {
    return this.errors;
  }

  /**
   * Checks if validation passed (no errors).
   * @returns {boolean} True if no errors
   */
  isValid() {
    return this.errors.length === 0;
  }
}

