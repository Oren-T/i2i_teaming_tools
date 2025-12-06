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
   * Uses the Active? column when present; falls back to legacy Permissions column.
   * @returns {string[]} Array of active staff names
   */
  getActiveStaffNames() {
    const names = [];
    const nameCol = this.getColumnIndex(DIRECTORY_COLUMNS.NAME);
    const activeCol = this.getColumnIndex(DIRECTORY_COLUMNS.ACTIVE);
    const permCol = this.getColumnIndex(DIRECTORY_COLUMNS.PERMISSIONS); // legacy

    if (nameCol === undefined) return names;

    for (let i = 1; i < this.data.length; i++) {
      const name = String(this.data[i][nameCol] || '').trim();
      if (!name) {
        continue;
      }

      if (!this.isRowActive(i, activeCol, permCol)) {
        continue;
      }

      names.push(name);
    }

    return names;
  }

  /**
   * Gets emails of active staff members.
   * Uses the Active? column when present; falls back to legacy Permissions column.
   * @returns {string[]} Array of active staff email addresses
   */
  getActiveStaffEmails() {
    const emails = [];
    const emailCol = this.getColumnIndex(DIRECTORY_COLUMNS.EMAIL);
    const activeCol = this.getColumnIndex(DIRECTORY_COLUMNS.ACTIVE);
    const permCol = this.getColumnIndex(DIRECTORY_COLUMNS.PERMISSIONS); // legacy

    if (emailCol === undefined) return emails;

    for (let i = 1; i < this.data.length; i++) {
      const email = String(this.data[i][emailCol] || '').trim();
      if (!email) {
        continue;
      }

      if (!this.isRowActive(i, activeCol, permCol)) {
        continue;
      }

      emails.push(email);
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

  /**
   * Determines whether a directory row should be considered "active" for dropdowns.
   * Prefers the Active? column when present; otherwise falls back to legacy Permissions.
   * @param {number} rowIndex - 1-based row index in the sheet (excluding header)
   * @param {number|undefined} activeCol - Active? column index (0-based) or undefined
   * @param {number|undefined} permCol - Legacy Permissions column index (0-based) or undefined
   * @returns {boolean} True if the row is active
   */
  isRowActive(rowIndex, activeCol, permCol) {
    const idx = rowIndex; // our this.data is 0-based with header at index 0

    // Prefer explicit Active? column when present
    if (activeCol !== undefined) {
      const flag = String(this.data[idx][activeCol] || '').trim().toLowerCase();
      // Only "yes" is treated as active; blank or "no" are inactive
      return flag === 'yes';
    }

    // Legacy behaviour: infer activity from Permissions column
    if (permCol !== undefined) {
      const perm = String(this.data[idx][permCol] || '').trim().toLowerCase();
      // Empty or explicit "none"/"no access" are inactive
      if (!perm || perm === 'none' || perm === 'no access') {
        return false;
      }
      return true;
    }

    // No explicit indicators – treat as active for compatibility
    return true;
  }

  /**
   * Gets structured access rows for permission evaluation.
   * Each row includes raw column values and normalized effective roles/scopes.
   * @returns {Object[]} Array of access row objects
   */
  getAccessRows() {
    const rows = [];

    const nameCol = this.getColumnIndex(DIRECTORY_COLUMNS.NAME);
    const emailCol = this.getColumnIndex(DIRECTORY_COLUMNS.EMAIL);
    const activeCol = this.getColumnIndex(DIRECTORY_COLUMNS.ACTIVE);
    const globalAccessCol = this.getColumnIndex(DIRECTORY_COLUMNS.GLOBAL_ACCESS);
    const mainFileRoleCol = this.getColumnIndex(DIRECTORY_COLUMNS.MAIN_FILE_ROLE);
    const projectScopeCol = this.getColumnIndex(DIRECTORY_COLUMNS.PROJECT_SCOPE);
    const legacyPermCol = this.getColumnIndex(DIRECTORY_COLUMNS.PERMISSIONS);

    if (emailCol === undefined) {
      return rows;
    }

    const hasNewPermissionColumns =
      globalAccessCol !== undefined ||
      mainFileRoleCol !== undefined ||
      projectScopeCol !== undefined;

    for (let i = 1; i < this.data.length; i++) {
      const row = this.data[i];
      const email = String(row[emailCol] || '').trim().toLowerCase();
      if (!email) {
        continue;
      }

      const name = nameCol !== undefined ? String(row[nameCol] || '').trim() : '';

      // Normalize Active? flag
      let activeFlag = 'blank';
      if (activeCol !== undefined) {
        const flag = String(row[activeCol] || '').trim().toLowerCase();
        if (flag === 'yes') {
          activeFlag = 'yes';
        } else if (flag === 'no') {
          activeFlag = 'no';
        } else {
          activeFlag = 'blank';
        }
      } else if (legacyPermCol !== undefined) {
        const perm = String(row[legacyPermCol] || '').trim().toLowerCase();
        if (!perm || perm === 'none' || perm === 'no access') {
          activeFlag = 'no';
        } else {
          activeFlag = 'yes';
        }
      } else {
        // No explicit indicators – treat as active for compatibility
        activeFlag = 'yes';
      }

      // Read raw permission values
      let globalAccessRaw = '';
      let mainFileRoleRaw = '';
      let projectScopeRaw = '';

      if (hasNewPermissionColumns) {
        if (globalAccessCol !== undefined) {
          globalAccessRaw = String(row[globalAccessCol] || '').trim();
        }
        if (mainFileRoleCol !== undefined) {
          mainFileRoleRaw = String(row[mainFileRoleCol] || '').trim();
        }
        if (projectScopeCol !== undefined) {
          projectScopeRaw = String(row[projectScopeCol] || '').trim();
        }
      } else if (legacyPermCol !== undefined) {
        // Legacy mapping: map single Permissions column to Main File Role only
        const perm = String(row[legacyPermCol] || '').trim().toLowerCase();
        if (perm === 'edit' || perm === 'editor') {
          mainFileRoleRaw = 'Editor';
        } else if (perm === 'view' || perm === 'viewer') {
          mainFileRoleRaw = 'Viewer';
        }
      }

      const {
        globalAccessRole,
        effectiveSpreadsheetRole,
        effectiveFolderScope
      } = computeEffectiveDirectoryAccess(globalAccessRaw, mainFileRoleRaw, projectScopeRaw);

      rows.push({
        rowIndex: i + 1, // 1-based for SpreadsheetApp APIs
        name,
        email,
        activeFlag,                  // 'yes' | 'no' | 'blank'
        globalAccessRaw,
        mainFileRoleRaw,
        projectScopeRaw,
        globalAccessRole,            // normalized: 'editor' | 'viewer' | 'none'
        effectiveSpreadsheetRole,    // DIRECTORY_ACCESS_ROLES.*
        effectiveFolderScope         // DIRECTORY_FOLDER_SCOPES.*
      });
    }

    return rows;
  }
}

/**
 * Computes normalized access values for a directory row based on
 * Global Access, Main File Role, and Project Scope.
 *
 * @param {string} globalAccessRaw - Raw Global Access cell value
 * @param {string} mainFileRoleRaw - Raw Main File Role cell value
 * @param {string} projectScopeRaw - Raw Project Scope cell value
 * @returns {Object} access object with normalized roles/scopes
 */
function computeEffectiveDirectoryAccess(globalAccessRaw, mainFileRoleRaw, projectScopeRaw) {
  const roles = DIRECTORY_ACCESS_ROLES;
  const scopes = DIRECTORY_FOLDER_SCOPES;

  const globalRole = normalizeDirectoryRole(globalAccessRaw);
  const mainRole = normalizeDirectoryRole(mainFileRoleRaw);
  const projectScope = normalizeProjectScope(projectScopeRaw);

  let effectiveSpreadsheetRole = roles.NONE;
  let effectiveFolderScope = scopes.ASSIGNED_ONLY;

  if (globalRole === roles.EDITOR) {
    // Global Editor: full editor on spreadsheet + all folders
    effectiveSpreadsheetRole = roles.EDITOR;
    effectiveFolderScope = scopes.ALL_EDITOR;
  } else if (globalRole === roles.VIEWER) {
    // Global Viewer: viewer on spreadsheet + all folders,
    // with explicit overrides allowed in the other columns.
    effectiveSpreadsheetRole = roles.VIEWER;
    effectiveFolderScope = scopes.ALL_VIEWER;

    if (mainRole === roles.EDITOR) {
      // Upgrade spreadsheet access only
      effectiveSpreadsheetRole = roles.EDITOR;
    }

    if (projectScope === scopes.ALL_EDITOR) {
      effectiveFolderScope = scopes.ALL_EDITOR;
    } else if (projectScope === scopes.ALL_VIEWER) {
      effectiveFolderScope = scopes.ALL_VIEWER;
    }
  } else {
    // No Global Access: spreadsheet and folders are controlled
    // purely by Main File Role and Project Scope.
    if (mainRole === roles.EDITOR) {
      effectiveSpreadsheetRole = roles.EDITOR;
    } else if (mainRole === roles.VIEWER) {
      effectiveSpreadsheetRole = roles.VIEWER;
    } else {
      effectiveSpreadsheetRole = roles.NONE;
    }

    if (projectScope === scopes.ALL_EDITOR || projectScope === scopes.ALL_VIEWER) {
      effectiveFolderScope = projectScope;
    } else {
      effectiveFolderScope = scopes.ASSIGNED_ONLY;
    }
  }

  return {
    globalAccessRole: globalRole,
    effectiveSpreadsheetRole,
    effectiveFolderScope
  };
}

/**
 * Normalizes a role-like string (Editor/Viewer/blank) to an internal value.
 * @param {string} value - Raw cell value
 * @returns {string} One of DIRECTORY_ACCESS_ROLES.*
 */
function normalizeDirectoryRole(value) {
  if (!value) {
    return DIRECTORY_ACCESS_ROLES.NONE;
  }

  const lower = String(value).trim().toLowerCase();
  if (lower === 'editor' || lower === 'edit') {
    return DIRECTORY_ACCESS_ROLES.EDITOR;
  }
  if (lower === 'viewer' || lower === 'view') {
    return DIRECTORY_ACCESS_ROLES.VIEWER;
  }

  return DIRECTORY_ACCESS_ROLES.NONE;
}

/**
 * Normalizes a Project Scope string to an internal scope value.
 * @param {string} value - Raw cell value
 * @returns {string} One of DIRECTORY_FOLDER_SCOPES.*
 */
function normalizeProjectScope(value) {
  if (!value) {
    return DIRECTORY_FOLDER_SCOPES.ASSIGNED_ONLY;
  }

  const lower = String(value).trim().toLowerCase();

  if (lower === 'all - editor' || lower === 'all-editor' || lower === 'all editor') {
    return DIRECTORY_FOLDER_SCOPES.ALL_EDITOR;
  }

  if (lower === 'all - viewer' || lower === 'all-viewer' || lower === 'all viewer') {
    return DIRECTORY_FOLDER_SCOPES.ALL_VIEWER;
  }

  return DIRECTORY_FOLDER_SCOPES.ASSIGNED_ONLY;
}

