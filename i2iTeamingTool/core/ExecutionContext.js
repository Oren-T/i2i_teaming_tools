/**
 * ExecutionContext class - Bundles all cached state for a single script execution.
 * Constructed once at the entry point and passed to all services.
 */
class ExecutionContext {
  /**
   * Creates a new ExecutionContext instance.
   * @param {string} spreadsheetId - The Main Projects File spreadsheet ID
   */
  constructor(spreadsheetId) {
    DEBUG && console.log(`ExecutionContext: Initializing for spreadsheet ${spreadsheetId}`);

    // Store the spreadsheet ID
    this.spreadsheetId = spreadsheetId;

    // Open the spreadsheet
    this.ss = SpreadsheetApp.openById(spreadsheetId);

    // Consistent timestamp for entire run
    this.now = new Date();

    // Initialize data layer
    this.initDataLayer();

    // Initialize services (after data layer is ready)
    this.initServices();

    DEBUG && console.log('ExecutionContext: Initialization complete');
  }

  /**
   * Initializes all data layer classes.
   */
  initDataLayer() {
    // Config - wraps the Config sheet
    const configSheet = this.ss.getSheetByName(SHEET_NAMES.CONFIG);
    if (!configSheet) {
      throw new Error(`Sheet "${SHEET_NAMES.CONFIG}" not found`);
    }
    this.config = new Config(configSheet);

    // Set DEBUG flag from config (do this early so subsequent logging respects it)
    this.initDebugMode();

    // IdAllocator - lock-protected ID generation
    this.idAllocator = new IdAllocator(this.config);

    // ProjectSheet - wraps the Projects sheet
    const projectsSheet = this.ss.getSheetByName(SHEET_NAMES.PROJECTS);
    if (!projectsSheet) {
      throw new Error(`Sheet "${SHEET_NAMES.PROJECTS}" not found`);
    }
    this.projectSheet = new ProjectSheet(projectsSheet);
    this.projectSheet.loadData();

    // SnapshotSheet - wraps the Status Snapshot sheet
    const snapshotSheet = this.ss.getSheetByName(SHEET_NAMES.STATUS_SNAPSHOT);
    if (!snapshotSheet) {
      throw new Error(`Sheet "${SHEET_NAMES.STATUS_SNAPSHOT}" not found`);
    }
    this.snapshotSheet = new SnapshotSheet(snapshotSheet);
    this.snapshotSheet.ensureHeaders();

    // Directory - wraps the Directory sheet
    const directorySheet = this.ss.getSheetByName(SHEET_NAMES.DIRECTORY);
    if (!directorySheet) {
      throw new Error(`Sheet "${SHEET_NAMES.DIRECTORY}" not found`);
    }
    this.directory = new Directory(directorySheet);

    // Codes - wraps the Codes sheet
    const codesSheet = this.ss.getSheetByName(SHEET_NAMES.CODES);
    if (!codesSheet) {
      throw new Error(`Sheet "${SHEET_NAMES.CODES}" not found`);
    }
    this.codes = new Codes(codesSheet);

    DEBUG && console.log('ExecutionContext: Data layer initialized');
  }

  /**
   * Initializes all service classes.
   */
  initServices() {
    // Validator - validates configuration and structure
    this.validator = new Validator(this.config, this.projectSheet);

    // NotificationService - email notifications
    this.notificationService = new NotificationService(this.config, this.directory);

    // ProjectService - project lifecycle operations
    this.projectService = new ProjectService(this);

    // MaintenanceService - daily maintenance tasks
    this.maintenanceService = new MaintenanceService(this);

    // FormService - Google Form dropdown sync
    this.formService = new FormService(this);

    // ValidationService - manages dropdown data validation rules
    this.validationService = new ValidationService(this);

    DEBUG && console.log('ExecutionContext: Services initialized');
  }

  /**
   * Gets the underlying Spreadsheet object.
   * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} The spreadsheet
   */
  getSpreadsheet() {
    return this.ss;
  }

  /**
   * Flushes all pending changes to the sheet.
   */
  flush() {
    this.projectSheet.flush();

    // Hide any rows marked for deletion
    this.projectService.hideDeletedRows();
  }

  /**
   * Runs validation on the context.
   * @throws {Error} If validation fails
   */
  validate() {
    this.validator.validate();
  }

  /**
   * Gets a summary of the context state.
   * @returns {Object} Context summary
   */
  getSummary() {
    return {
      spreadsheetId: this.spreadsheetId,
      timestamp: this.now.toISOString(),
      config: {
        districtId: this.config.districtId,
        schoolYear: this.config.schoolYear,
        nextSerial: this.config.nextSerial,
        debugMode: DEBUG
      },
      projectCount: this.projectSheet.getProjects().length,
      readyCount: this.projectSheet.getReadyProjects().length,
      createdCount: this.projectSheet.getCreatedProjects().length,
      staffCount: this.directory.getAllStaffNames().length
    };
  }

  /**
   * Initializes the DEBUG flag from the Config sheet.
   * Looks for "Debug Mode" key with value "true" (case-insensitive).
   */
  initDebugMode() {
    const debugValue = this.config.get('Debug Mode');
    if (debugValue !== undefined && debugValue !== null) {
      const debugStr = String(debugValue).trim().toLowerCase();
      DEBUG = (debugStr === 'true' || debugStr === 'yes' || debugStr === '1');
    }

    if (DEBUG) {
      console.log('ExecutionContext: Debug mode ENABLED via Config sheet');
    }
  }
}

