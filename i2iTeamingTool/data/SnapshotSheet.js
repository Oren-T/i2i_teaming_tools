/**
 * SnapshotSheet class - Manages the Status Snapshot sheet.
 * Tracks previous day's project statuses for change detection in daily maintenance.
 */
class SnapshotSheet {
  /**
   * Creates a new SnapshotSheet instance.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Status Snapshot sheet
   */
  constructor(sheet) {
    this.sheet = sheet;
    this.data = null;
    this.PROJECT_ID_COL = 0;
    this.STATUS_COL = 1;
  }

  /**
   * Loads the snapshot data into memory.
   * @returns {Map<string, string>} Map of projectId -> projectStatus
   */
  loadSnapshot() {
    if (!this.sheet) {
      throw new Error('Status Snapshot sheet not found');
    }

    const dataRange = this.sheet.getDataRange();
    this.data = dataRange.getValues();

    const snapshot = new Map();

    // Skip header row (Row 1)
    for (let i = 1; i < this.data.length; i++) {
      const projectId = String(this.data[i][this.PROJECT_ID_COL] || '').trim();
      const status = String(this.data[i][this.STATUS_COL] || '').trim();

      if (projectId) {
        snapshot.set(projectId, status);
      }
    }

    DEBUG && console.log(`SnapshotSheet: Loaded ${snapshot.size} entries from snapshot`);
    return snapshot;
  }

  /**
   * Overwrites the snapshot with current project statuses.
   * Clears existing data and writes new snapshot.
   * @param {Map<string, string>} currentStatuses - Map of projectId -> projectStatus
   */
  overwriteWithCurrent(currentStatuses) {
    if (!this.sheet) {
      throw new Error('Status Snapshot sheet not found');
    }

    // Clear existing data (except header)
    const lastRow = this.sheet.getLastRow();
    if (lastRow > 1) {
      this.sheet.deleteRows(2, lastRow - 1);
    }

    // Prepare new data
    const newData = [];
    for (const [projectId, status] of currentStatuses) {
      if (projectId) {
        newData.push([projectId, status]);
      }
    }

    // Write new data
    if (newData.length > 0) {
      const range = this.sheet.getRange(2, 1, newData.length, 2);
      range.setValues(newData);
    }

    DEBUG && console.log(`SnapshotSheet: Wrote ${newData.length} entries to snapshot`);
  }

  /**
   * Detects status changes between previous snapshot and current statuses.
   * @param {Map<string, string>} currentStatuses - Current project statuses
   * @returns {Object[]} Array of change objects {projectId, oldStatus, newStatus}
   */
  detectChanges(currentStatuses) {
    const previousSnapshot = this.loadSnapshot();
    const changes = [];

    // Check for changed and new statuses
    for (const [projectId, newStatus] of currentStatuses) {
      const oldStatus = previousSnapshot.get(projectId);

      if (oldStatus === undefined) {
        // New project - not a "change" for notification purposes
        // (it was never in the snapshot before)
        DEBUG && console.log(`SnapshotSheet: New project detected: ${projectId}`);
      } else if (oldStatus !== newStatus) {
        changes.push({
          projectId,
          oldStatus,
          newStatus
        });
        DEBUG && console.log(`SnapshotSheet: Status change: ${projectId} "${oldStatus}" -> "${newStatus}"`);
      }
    }

    DEBUG && console.log(`SnapshotSheet: Detected ${changes.length} status change(s)`);
    return changes;
  }

  /**
   * Gets the count of entries in the snapshot.
   * @returns {number} Number of snapshot entries
   */
  getEntryCount() {
    if (!this.data) {
      this.loadSnapshot();
    }
    return Math.max(0, (this.data ? this.data.length : 1) - 1);
  }

  /**
   * Checks if the snapshot is empty (has no project entries).
   * @returns {boolean} True if empty
   */
  isEmpty() {
    return this.getEntryCount() === 0;
  }

  /**
   * Initializes the snapshot from current project statuses if empty.
   * Should be called on first run.
   * @param {Map<string, string>} currentStatuses - Current project statuses
   * @returns {boolean} True if initialization was performed
   */
  initializeIfEmpty(currentStatuses) {
    if (this.isEmpty()) {
      DEBUG && console.log('SnapshotSheet: Initializing empty snapshot');
      this.overwriteWithCurrent(currentStatuses);
      return true;
    }
    return false;
  }

  /**
   * Ensures the snapshot header row exists.
   */
  ensureHeaders() {
    if (!this.sheet) return;

    const headerRange = this.sheet.getRange(1, 1, 1, 2);
    const headers = headerRange.getValues()[0];

    if (headers[0] !== 'project_id' || headers[1] !== 'project_status') {
      headerRange.setValues([['project_id', 'project_status']]);
      DEBUG && console.log('SnapshotSheet: Set header row');
    }
  }

  /**
   * Gets the underlying Sheet object.
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet
   */
  getSheet() {
    return this.sheet;
  }
}

