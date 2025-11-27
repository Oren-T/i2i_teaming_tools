/**
 * Project class - Row-level abstraction with typed accessors and dirty tracking.
 * Wraps a single project row from the Projects sheet.
 */
class Project {
  /**
   * Creates a new Project instance.
   * @param {Array} rowData - The raw row data array
   * @param {number} rowIndex - The 1-based row index in the sheet
   * @param {Map} columnMap - Map of column key -> column index (0-based)
   */
  constructor(rowData, rowIndex, columnMap) {
    this.originalData = [...rowData];
    this.currentData = [...rowData];
    this.rowIndex = rowIndex;
    this.columnMap = columnMap;
    this.dirtyColumns = new Set();
  }

  // ===== CORE ACCESSORS =====

  /**
   * Gets a value by column key.
   * @param {string} key - The column key (e.g., 'project_id')
   * @returns {*} The cell value
   */
  get(key) {
    const colIndex = this.columnMap.get(key);
    if (colIndex === undefined) {
      DEBUG && console.log(`Project: Unknown column key "${key}"`);
      return undefined;
    }
    return this.currentData[colIndex];
  }

  /**
   * Sets a value by column key and marks it as dirty.
   * @param {string} key - The column key
   * @param {*} value - The value to set
   */
  set(key, value) {
    const colIndex = this.columnMap.get(key);
    if (colIndex === undefined) {
      DEBUG && console.log(`Project: Unknown column key "${key}"`);
      return;
    }

    // Only mark dirty if value actually changed
    if (this.currentData[colIndex] !== value) {
      this.currentData[colIndex] = value;
      this.dirtyColumns.add(key);
    }
  }

  /**
   * Checks if any values have been modified.
   * @returns {boolean} True if dirty
   */
  isDirty() {
    return this.dirtyColumns.size > 0;
  }

  /**
   * Gets the list of modified column keys.
   * @returns {string[]} Array of dirty column keys
   */
  getDirtyKeys() {
    return Array.from(this.dirtyColumns);
  }

  /**
   * Gets dirty entries as {key, colIndex, value} objects.
   * @returns {Object[]} Array of dirty entry objects
   */
  getDirtyEntries() {
    const entries = [];
    for (const key of this.dirtyColumns) {
      const colIndex = this.columnMap.get(key);
      entries.push({
        key,
        colIndex,
        value: this.currentData[colIndex]
      });
    }
    return entries;
  }

  /**
   * Clears the dirty state (call after flushing to sheet).
   */
  clearDirty() {
    this.originalData = [...this.currentData];
    this.dirtyColumns.clear();
  }

  /**
   * Gets the 1-based row index.
   * @returns {number} Row index
   */
  getRowIndex() {
    return this.rowIndex;
  }

  // ===== TYPED GETTERS =====

  /**
   * @returns {string} Project ID (e.g., "NUSD-25_26-0024")
   */
  get projectId() {
    return String(this.get('project_id') || '').trim();
  }

  /**
   * @returns {Date|null} Created at timestamp
   */
  get createdAt() {
    return parseDate(this.get('created_at'));
  }

  /**
   * @returns {string} School year (e.g., "25_26")
   */
  get schoolYear() {
    return String(this.get('school_year') || '').trim();
  }

  /**
   * @returns {string} Goal number
   */
  get goalNumber() {
    return String(this.get('goal_number') || '').trim();
  }

  /**
   * @returns {string} Action number
   */
  get actionNumber() {
    return String(this.get('action_number') || '').trim();
  }

  /**
   * @returns {string} Category (e.g., "LCAP")
   */
  get category() {
    return String(this.get('category') || '').trim();
  }

  /**
   * @returns {string} Project name/title
   */
  get projectName() {
    return String(this.get('project_name') || '').trim();
  }

  /**
   * @returns {string} Project description
   */
  get description() {
    return String(this.get('description') || '').trim();
  }

  /**
   * @returns {string} Assignee(s) - may be comma-separated
   */
  get assignee() {
    return String(this.get('assignee') || '').trim();
  }

  /**
   * @returns {string[]} Array of assignee names/emails
   */
  get assignees() {
    return parseCommaSeparated(this.assignee);
  }

  /**
   * @returns {string} Requested by (person who submitted)
   */
  get requestedBy() {
    return String(this.get('requested_by') || '').trim();
  }

  /**
   * @returns {Date|null} Due date
   */
  get dueDate() {
    return parseDate(this.get('due_date'));
  }

  /**
   * @returns {string} Project status (e.g., "On track", "Complete")
   */
  get projectStatus() {
    return String(this.get('project_status') || '').trim();
  }

  /**
   * @returns {Date|null} Completed at timestamp
   */
  get completedAt() {
    return parseDate(this.get('completed_at'));
  }

  /**
   * @returns {string} Raw reminder offsets string
   */
  get reminderOffsetsRaw() {
    return String(this.get('reminder_offsets') || '').trim();
  }

  /**
   * @returns {number[]} Parsed reminder offsets as integers
   */
  get reminderOffsets() {
    return parseReminderOffsets(this.reminderOffsetsRaw);
  }

  /**
   * @returns {string} Automation status (e.g., "Ready", "Created")
   */
  get automationStatus() {
    return String(this.get('automation_status') || '').trim();
  }

  /**
   * @returns {string} Calendar event ID
   */
  get calendarEventId() {
    return String(this.get('calendar_event_id') || '').trim();
  }

  /**
   * @returns {string} Folder ID
   */
  get folderId() {
    return String(this.get('folder_id') || '').trim();
  }

  /**
   * @returns {string} Notes
   */
  get notes() {
    return String(this.get('notes') || '').trim();
  }

  // ===== TYPED SETTERS =====

  set projectId(value) {
    this.set('project_id', value);
  }

  set createdAt(value) {
    this.set('created_at', value);
  }

  set schoolYear(value) {
    this.set('school_year', value);
  }

  set projectStatus(value) {
    this.set('project_status', value);
  }

  set completedAt(value) {
    this.set('completed_at', value);
  }

  set automationStatus(value) {
    this.set('automation_status', value);
  }

  set calendarEventId(value) {
    this.set('calendar_event_id', value);
  }

  set folderId(value) {
    this.set('folder_id', value);
  }

  /**
   * @returns {string} File ID (the copied project spreadsheet)
   */
  get fileId() {
    return String(this.get('file_id') || '').trim();
  }

  set fileId(value) {
    this.set('file_id', value);
  }

  // ===== COMPUTED PROPERTIES =====

  /**
   * Gets the project folder URL.
   * @returns {string} Google Drive folder URL
   */
  get folderUrl() {
    return folderIdToUrl(this.folderId);
  }

  /**
   * Gets the display title in format "[ID] Name".
   * @returns {string} Display title
   */
  get displayTitle() {
    return formatProjectTitle(this.projectId, this.projectName);
  }

  /**
   * Calculates days until due from a reference date.
   * @param {Date} refDate - Reference date (default: today)
   * @returns {number} Days until due (negative if past due)
   */
  daysUntilDue(refDate = null) {
    const from = refDate || getStartOfToday();
    const to = this.dueDate;
    if (!to) return 0;
    return daysBetween(from, to);
  }

  // ===== PREDICATES =====

  /**
   * @returns {boolean} True if automation status is Ready
   */
  get isReady() {
    return this.automationStatus === AUTOMATION_STATUS.READY;
  }

  /**
   * @returns {boolean} True if automation status is Created
   */
  get isCreated() {
    return this.automationStatus === AUTOMATION_STATUS.CREATED;
  }

  /**
   * @returns {boolean} True if automation status is Updated
   */
  get isUpdated() {
    return this.automationStatus === AUTOMATION_STATUS.UPDATED;
  }

  /**
   * @returns {boolean} True if automation status is Error
   */
  get isError() {
    return this.automationStatus === AUTOMATION_STATUS.ERROR;
  }

  /**
   * @returns {boolean} True if pending delete (either notify or don't notify)
   */
  get isPendingDelete() {
    return this.automationStatus === AUTOMATION_STATUS.DELETE_NOTIFY ||
           this.automationStatus === AUTOMATION_STATUS.DELETE_NO_NOTIFY;
  }

  /**
   * @returns {boolean} True if delete should notify attendees
   */
  get shouldNotifyOnDelete() {
    return this.automationStatus === AUTOMATION_STATUS.DELETE_NOTIFY;
  }

  /**
   * @returns {boolean} True if automation status is Deleted
   */
  get isDeleted() {
    return this.automationStatus === AUTOMATION_STATUS.DELETED;
  }

  /**
   * @returns {boolean} True if project status is Complete
   */
  get isComplete() {
    return this.projectStatus === PROJECT_STATUS.COMPLETE;
  }

  /**
   * @returns {boolean} True if project status is Late
   */
  get isLate() {
    return this.projectStatus === PROJECT_STATUS.LATE;
  }

  /**
   * @returns {boolean} True if row has a project ID assigned
   */
  get hasProjectId() {
    return this.projectId !== '';
  }

  /**
   * @returns {boolean} True if the row appears to be empty/blank
   */
  get isBlankRow() {
    return !this.projectName && !this.projectId && !this.automationStatus;
  }

  /**
   * Checks if a reminder should be sent for a given offset on a reference date.
   * @param {number} offset - Days before due
   * @param {Date} refDate - Reference date (default: today)
   * @returns {boolean} True if reminder should be sent
   */
  isDueForReminder(offset, refDate = null) {
    const daysUntil = this.daysUntilDue(refDate);
    return daysUntil === offset;
  }

  /**
   * @returns {boolean} True if due date is today
   */
  isDueToday(refDate = null) {
    const due = this.dueDate;
    const today = refDate || getStartOfToday();
    return due && isSameDay(due, today);
  }

  /**
   * @returns {boolean} True if due date has passed
   */
  isPastDue(refDate = null) {
    return this.daysUntilDue(refDate) < 0;
  }

  // ===== UTILITY METHODS =====

  /**
   * Gets all relevant emails for this project (assignees + requested by).
   * @param {Directory} directory - Directory for name-to-email resolution
   * @returns {string[]} Array of email addresses
   */
  getAllRecipientEmails(directory) {
    const emails = new Set();

    // Add assignee emails
    for (const assignee of this.assignees) {
      const email = directory.resolveToEmail(assignee);
      if (email) emails.add(email);
    }

    // Add requested by email
    const requestedByEmail = directory.resolveToEmail(this.requestedBy);
    if (requestedByEmail) emails.add(requestedByEmail);

    return Array.from(emails);
  }

  /**
   * Gets only the assignee emails.
   * @param {Directory} directory - Directory for name-to-email resolution
   * @returns {string[]} Array of assignee email addresses
   */
  getAssigneeEmails(directory) {
    return directory.resolveAllToEmails(this.assignees);
  }

  /**
   * Creates a token values object for email template substitution.
   * @param {Directory} directory - Directory for name lookups
   * @param {Object} extras - Additional token values to merge
   * @returns {Object} Token values object
   */
  getTokenValues(directory, extras = {}) {
    const assigneeNames = this.assignees.map(a => {
      const name = directory.getNameByEmail(a);
      return name || a;
    });

    const requestedByName = directory.getNameByEmail(this.requestedBy) || this.requestedBy;

    return {
      PROJECT_TITLE: this.projectName,
      PROJECT_ID: this.projectId,
      ASSIGNEE_NAME: joinCommaSeparated(assigneeNames),
      REQUESTED_BY_NAME: requestedByName,
      CATEGORY: this.category,
      DEADLINE: formatDate(this.dueDate),
      DESCRIPTION: this.description,
      FOLDER_LINK: this.folderUrl,
      NEW_STATUS: this.projectStatus,
      ...extras
    };
  }

  /**
   * Returns a debug string representation.
   * @returns {string} Debug string
   */
  toString() {
    return `Project[${this.rowIndex}]: ${this.displayTitle} (status: ${this.automationStatus})`;
  }
}

