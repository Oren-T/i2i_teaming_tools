/**
 * ProjectSheet class - Manages the main Projects sheet.
 * Handles column indexing from Row 2, creates Project instances, and batch writes.
 */
class ProjectSheet {
  /**
   * Creates a new ProjectSheet instance.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Projects sheet
   */
  constructor(sheet) {
    this.sheet = sheet;
    this.data = null;
    this.columnMap = null;
    this.projects = [];
    this.headerRowCount = 2; // Row 1: labels, Row 2: keys
  }

  /**
   * Loads all data from the Projects sheet and creates Project instances.
   */
  loadData() {
    if (!this.sheet) {
      throw new Error('Projects sheet not found');
    }

    const dataRange = this.sheet.getDataRange();
    this.data = dataRange.getValues();

    if (this.data.length < this.headerRowCount) {
      DEBUG && console.log('ProjectSheet: Sheet has fewer than 2 header rows');
      this.columnMap = new Map();
      this.projects = [];
      return;
    }

    // Build column map from Row 2 (index 1)
    this.columnMap = new Map();
    const keyRow = this.data[1];
    for (let i = 0; i < keyRow.length; i++) {
      const key = String(keyRow[i]).trim();
      if (key) {
        this.columnMap.set(key, i);
      }
    }

    // Create Project instances for data rows (starting at Row 3 = index 2)
    this.projects = [];
    for (let i = this.headerRowCount; i < this.data.length; i++) {
      const rowData = this.data[i];
      const rowIndex = i + 1; // 1-based row number
      const project = new Project(rowData, rowIndex, this.columnMap);

      // Skip completely blank rows
      if (!project.isBlankRow) {
        this.projects.push(project);
      }
    }

    DEBUG && console.log(`ProjectSheet: Loaded ${this.projects.length} projects with ${this.columnMap.size} columns`);
  }

  /**
   * Gets the column keys array (Row 2 values).
   * @returns {string[]} Array of column keys
   */
  getColumnKeys() {
    if (!this.data || this.data.length < 2) {
      return [];
    }
    return this.data[1].map(k => String(k).trim()).filter(k => k !== '');
  }

  /**
   * Gets the column index for an internal key.
   * @param {string} key - The column key (e.g., 'project_id')
   * @returns {number|undefined} 0-based column index
   */
  getColumnIndex(key) {
    return this.columnMap ? this.columnMap.get(key) : undefined;
  }

  /**
   * Gets all Project instances.
   * @returns {Project[]} Array of projects
   */
  getProjects() {
    return this.projects;
  }

  /**
   * Gets projects matching a predicate.
   * @param {Function} predicate - Filter function (project) => boolean
   * @returns {Project[]} Filtered projects
   */
  getProjectsWhere(predicate) {
    return this.projects.filter(predicate);
  }

  /**
   * Gets projects with automation_status = 'Ready'.
   * @returns {Project[]} Ready projects
   */
  getReadyProjects() {
    return this.getProjectsWhere(p => p.isReady);
  }

  /**
   * Gets projects with automation_status = 'Updated'.
   * @returns {Project[]} Updated projects
   */
  getUpdatedProjects() {
    return this.getProjectsWhere(p => p.isUpdated);
  }

  /**
   * Gets projects with automation_status = 'Delete (Notify)' or 'Delete (Don't Notify)'.
   * @returns {Project[]} Pending delete projects
   */
  getPendingDeleteProjects() {
    return this.getProjectsWhere(p => p.isPendingDelete);
  }

  /**
   * Gets projects with automation_status = 'Created' (active projects).
   * @returns {Project[]} Created projects
   */
  getCreatedProjects() {
    return this.getProjectsWhere(p => p.isCreated);
  }

  /**
   * Gets projects with automation_status = 'Error'.
   * @returns {Project[]} Error projects
   */
  getErrorProjects() {
    return this.getProjectsWhere(p => p.isError);
  }

  /**
   * Gets active (non-deleted) projects.
   * @returns {Project[]} Active projects
   */
  getActiveProjects() {
    return this.getProjectsWhere(p => !p.isDeleted);
  }

  /**
   * Gets incomplete projects (not completed).
   * @returns {Project[]} Incomplete projects
   */
  getIncompleteProjects() {
    return this.getProjectsWhere(p => !p.isComplete && !p.isDeleted);
  }

  /**
   * Finds a project by ID.
   * @param {string} projectId - The project ID to find
   * @returns {Project|undefined} The project or undefined
   */
  findByProjectId(projectId) {
    if (!projectId) return undefined;
    return this.projects.find(p => p.projectId === projectId);
  }

  /**
   * Gets the column map for constructing new Project instances.
   * @returns {Map} Column key -> index map
   */
  getColumnMap() {
    return this.columnMap;
  }

  /**
   * Flushes all dirty projects back to the sheet.
   * Only writes modified cells for efficiency.
   */
  flush() {
    const dirtyProjects = this.projects.filter(p => p.isDirty());

    if (dirtyProjects.length === 0) {
      DEBUG && console.log('ProjectSheet.flush: No dirty projects');
      return;
    }

    DEBUG && console.log(`ProjectSheet.flush: Writing ${dirtyProjects.length} dirty project(s)`);

    for (const project of dirtyProjects) {
      const entries = project.getDirtyEntries();

      for (const { colIndex, value } of entries) {
        // colIndex is 0-based, need to convert to 1-based for getRange
        const range = this.sheet.getRange(project.getRowIndex(), colIndex + 1);
        range.setValue(value);
      }

      project.clearDirty();

      DEBUG && console.log(`ProjectSheet.flush: Updated row ${project.getRowIndex()}: ${project.getDirtyKeys().join(', ')}`);
    }

    // Force pending writes to the spreadsheet immediately.
    // This ensures changes are visible before any subsequent reads or API calls.
    SpreadsheetApp.flush();
  }

  /**
   * Appends a new row to the sheet with the given data.
   * @param {Object} data - Key-value pairs for the new row
   * @returns {Project} The created Project instance
   */
  appendRow(data) {
    // Build row array based on column map
    const maxCol = Math.max(...this.columnMap.values()) + 1;
    const rowArray = new Array(maxCol).fill('');

    for (const [key, value] of Object.entries(data)) {
      const colIndex = this.columnMap.get(key);
      if (colIndex !== undefined) {
        rowArray[colIndex] = value;
      }
    }

    // Append to sheet
    this.sheet.appendRow(rowArray);

    // Get the new row number
    const newRowIndex = this.sheet.getLastRow();

    // Create Project instance
    const project = new Project(rowArray, newRowIndex, this.columnMap);
    this.projects.push(project);

    DEBUG && console.log(`ProjectSheet.appendRow: Added row ${newRowIndex}`);

    return project;
  }

  /**
   * Hides a row (used for soft-delete).
   * @param {Project} project - The project to hide
   */
  hideRow(project) {
    const rowIndex = project.getRowIndex();
    this.sheet.hideRows(rowIndex);
    DEBUG && console.log(`ProjectSheet.hideRow: Hidden row ${rowIndex}`);
  }

  /**
   * Gets the underlying Sheet object.
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet
   */
  getSheet() {
    return this.sheet;
  }

  /**
   * Refreshes data from the sheet (re-loads).
   */
  refresh() {
    this.loadData();
  }

  /**
   * Gets project statuses as a Map for snapshot comparison.
   * Only includes active (non-deleted) projects.
   * @returns {Map<string, string>} Map of projectId -> projectStatus
   */
  getStatusMap() {
    const map = new Map();
    for (const project of this.projects) {
      // Only include active projects (exclude deleted/hidden)
      if (project.projectId && !project.isDeleted) {
        map.set(project.projectId, project.projectStatus);
      }
    }
    return map;
  }
}

