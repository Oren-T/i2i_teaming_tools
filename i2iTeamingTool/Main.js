/**
 * Main.js - Public API for the i2i Teaming Tool Library.
 *
 * These functions are exposed to client scripts and handle all business logic.
 * Client scripts should call these functions with their spreadsheet ID.
 *
 * Library identifier: i2iTT
 * Client usage: i2iTT.processNewProjects(SPREADSHEET_ID)
 */

// ===== PUBLIC API FUNCTIONS =====

/**
 * Processes projects with Ready, Updated, or Delete status.
 * Called by the 10-minute batch trigger or manual "Run Now" action.
 *
 * @param {string} spreadsheetId - The Main Projects File spreadsheet ID
 * @throws {Error} If validation fails or processing encounters an unrecoverable error
 */
function processNewProjects(spreadsheetId) {
  console.log('=== processNewProjects starting ===');

  // Acquire script lock to prevent overlapping runs
  const lock = LockService.getScriptLock();
  // Wait up to 30 seconds for the lock
  const acquired = lock.tryLock(30000);

  if (!acquired) {
    console.log('processNewProjects: Could not acquire lock, another instance may be running');
    return;
  }

  let ctx;

  try {
    ctx = new ExecutionContext(spreadsheetId);
    ctx.validate();

    // Process Ready projects (create folder, templates, calendar event)
    ctx.projectService.processReadyProjects();

    // Process Updated projects (re-sync calendar event)
    ctx.projectService.processUpdatedProjects();

    // Process Delete requests (cancel event, hide row)
    ctx.projectService.processDeleteRequests();

    // Flush all changes to the sheet
    ctx.flush();

    console.log('=== processNewProjects completed ===');

  } catch (error) {
    console.error(`processNewProjects error: ${error.message}`);

    const messageLines = [
      `Error: ${error.message}`,
      '',
      'Function: processNewProjects',
      `Spreadsheet ID: ${spreadsheetId}`,
      '',
      `Stack: ${error.stack || 'N/A'}`
    ];
    sendAdminErrorNotification(
      spreadsheetId,
      ctx,
      'Project Processing Failed',
      messageLines.join('\n')
    );

    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Runs daily maintenance tasks.
 * Called by the daily 8am trigger.
 *
 * @param {string} spreadsheetId - The Main Projects File spreadsheet ID
 * @throws {Error} If validation fails or maintenance tasks encounter an unrecoverable error
 */
function runDailyMaintenance(spreadsheetId) {
  console.log('=== runDailyMaintenance starting ===');

  // Acquire script lock to prevent overlapping runs (e.g. manual + trigger)
  const lock = LockService.getScriptLock();
  // Wait up to 30 seconds for the lock
  const acquired = lock.tryLock(30000);

  if (!acquired) {
    console.log('runDailyMaintenance: Could not acquire lock, another instance may be running');
    return;
  }

  let ctx;

  try {
    ctx = new ExecutionContext(spreadsheetId);
    ctx.validate();

    // Run all daily maintenance tasks
    ctx.maintenanceService.runDailyMaintenance();

    // Flush any changes
    ctx.flush();

    console.log('=== runDailyMaintenance completed ===');

  } catch (error) {
    console.error(`runDailyMaintenance error: ${error.message}`);

    const messageLines = [
      `Error: ${error.message}`,
      '',
      'Function: runDailyMaintenance',
      `Spreadsheet ID: ${spreadsheetId}`,
      '',
      `Stack: ${error.stack || 'N/A'}`
    ];
    sendAdminErrorNotification(
      spreadsheetId,
      ctx,
      'Daily Maintenance Failed',
      messageLines.join('\n')
    );

    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Syncs Google Form dropdowns with Directory and Codes sheets.
 * Called manually from the menu or when Directory/Codes are edited.
 *
 * @param {string} spreadsheetId - The Main Projects File spreadsheet ID
 * @throws {Error} If form sync fails
 */
function syncFormDropdowns(spreadsheetId) {
  console.log('=== syncFormDropdowns starting ===');

  let ctx;

  try {
    ctx = new ExecutionContext(spreadsheetId);

    // Sync all form dropdowns
    ctx.formService.syncAllDropdowns();

    console.log('=== syncFormDropdowns completed ===');

  } catch (error) {
    console.error(`syncFormDropdowns error: ${error.message}`);

    const messageLines = [
      `Error: ${error.message}`,
      '',
      'Function: syncFormDropdowns',
      `Spreadsheet ID: ${spreadsheetId}`,
      '',
      `Stack: ${error.stack || 'N/A'}`
    ];
    sendAdminErrorNotification(
      spreadsheetId,
      ctx,
      'Sync Form Dropdowns Failed',
      messageLines.join('\n')
    );

    throw error;
  }
}

/**
 * Handles a form submission event.
 * Called by the onFormSubmit trigger (spreadsheet-bound).
 *
 * @param {string} spreadsheetId - The Main Projects File spreadsheet ID
 * @param {Object} event - The form submission event object
 * @throws {Error} If validation fails or form processing encounters an error
 */
function handleFormSubmission(spreadsheetId, event) {
  console.log('=== handleFormSubmission starting ===');

  // Acquire script lock
  const lock = LockService.getScriptLock();
  // Wait up to 5 minutes (300,000 ms) for the lock
  const acquired = lock.tryLock(300000);

  if (!acquired) {
    const errorMsg = 'handleFormSubmission: Could not acquire lock after 5 minutes. Submission processing FAILED.';
    console.error(errorMsg);

    // Attempt to send error email notification (best-effort, even if full context cannot be initialized)
    sendAdminErrorNotification(
      spreadsheetId,
      null,
      'Form Submission Lock Timeout',
      errorMsg
    );

    // Throw error so the trigger is marked as failed in Apps Script dashboard
    throw new Error(errorMsg);
  }

  let ctx;

  try {
    ctx = new ExecutionContext(spreadsheetId);
    ctx.validate();

    // Normalize form response and append to Projects sheet
    ctx.projectService.normalizeAndAppendFormResponse(event);

    // Process the newly added Ready project immediately
    ctx.projectService.processReadyProjects();

    // Flush all changes
    ctx.flush();

    console.log('=== handleFormSubmission completed ===');

  } catch (error) {
    console.error(`handleFormSubmission error: ${error.message}`);

    const messageLines = [
      `Error: ${error.message}`,
      '',
      'Function: handleFormSubmission',
      `Spreadsheet ID: ${spreadsheetId}`,
      '',
      `Stack: ${error.stack || 'N/A'}`
    ];
    sendAdminErrorNotification(
      spreadsheetId,
      ctx,
      'Form Submission Failed',
      messageLines.join('\n')
    );

    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Handles edit events on the Main Projects File.
 * Called by the onEdit trigger.
 *
 * @param {string} spreadsheetId - The Main Projects File spreadsheet ID
 * @param {Object} event - The edit event object
 */
function handleEdit(spreadsheetId, event) {
  const range = event.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  DEBUG && console.log(`handleEdit: Sheet "${sheetName}", cell ${range.getA1Notation()}`);

  try {
    const ctx = new ExecutionContext(spreadsheetId);

    // Handle edits to the Projects sheet
    if (sheetName === SHEET_NAMES.PROJECTS) {
      handleProjectsEdit(ctx, event);
    }

    // Handle edits to the Directory sheet
    if (sheetName === SHEET_NAMES.DIRECTORY) {
      DEBUG && console.log('handleEdit: Directory edited, syncing form dropdowns');
      ctx.formService.syncAssigneeDropdown();
    }

    // Handle edits to the Codes sheet
    if (sheetName === SHEET_NAMES.CODES) {
      DEBUG && console.log('handleEdit: Codes edited, syncing form dropdowns');
      ctx.formService.syncCategoryDropdown();
    }

    // Flush any changes
    ctx.flush();

  } catch (error) {
    console.error(`handleEdit error: ${error.message}`);
    // Don't throw - onEdit triggers should fail silently
  }
}

/**
 * Handles edit events specifically for the Projects sheet.
 * Sets completed_at timestamp when project_status changes to "Complete".
 *
 * @param {ExecutionContext} ctx - The execution context
 * @param {Object} event - The edit event object
 */
function handleProjectsEdit(ctx, event) {
  const range = event.range;
  const sheet = range.getSheet();

  // Get the column that was edited
  const col = range.getColumn();
  const row = range.getRow();

  // Skip header rows
  if (row <= 2) {
    return;
  }

  // Find the project_status column index
  const statusColIndex = ctx.projectSheet.getColumnIndex('project_status');
  if (statusColIndex === undefined) {
    return;
  }

  // Check if project_status column was edited (convert to 1-based)
  if (col === statusColIndex + 1) {
    const newValue = String(event.value || '').trim();
    const oldValue = String(event.oldValue || '').trim();

    DEBUG && console.log(`handleProjectsEdit: Status changed from "${oldValue}" to "${newValue}"`);

    // If status changed to Complete (and wasn't already)
    if (newValue === PROJECT_STATUS.COMPLETE && oldValue !== PROJECT_STATUS.COMPLETE) {
      const completedAtColIndex = ctx.projectSheet.getColumnIndex('completed_at');
      if (completedAtColIndex !== undefined) {
        sheet.getRange(row, completedAtColIndex + 1).setValue(new Date());
        console.log(`Set completed_at for row ${row}`);
      }
    }
  }
}

/**
 * Refreshes sharing permissions on the Main Projects File based on Directory roles.
 * Syncs permissions for the main spreadsheet, root folder, Project Folders parent,
 * and all project folders using the Directory permission model.
 * Called manually from the menu.
 *
 * @param {string} spreadsheetId - The Main Projects File spreadsheet ID
 */
function refreshPermissions(spreadsheetId) {
  console.log('=== refreshPermissions starting ===');

  let ctx;

  try {
    ctx = new ExecutionContext(spreadsheetId);
    ctx.permissionService.refreshAllPermissions();

    console.log('=== refreshPermissions completed ===');

  } catch (error) {
    console.error(`refreshPermissions error: ${error.message}`);

    const messageLines = [
      `Error: ${error.message}`,
      '',
      'Function: refreshPermissions',
      `Spreadsheet ID: ${spreadsheetId}`,
      '',
      `Stack: ${error.stack || 'N/A'}`
    ];
    sendAdminErrorNotification(
      spreadsheetId,
      ctx,
      'Permission Refresh Failed',
      messageLines.join('\n')
    );

    throw error;
  }
}

/**
 * Creates the custom menu in the spreadsheet UI.
 * Called by the onOpen trigger.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sSht - The spreadsheet (optional, uses active if not provided)
 */
function createMenu(sSht) {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Teaming Tool')
    .addItem('Run Now', 'manualRunNow')
    .addItem('Sync Form Dropdowns', 'manualSyncDropdowns')
    .addItem('Refresh Permissions', 'manualRefreshPermissions')
    .addSeparator()
    .addItem('View Status Summary', 'showStatusSummary')
    .addSeparator()
    .addSubMenu(ui.createMenu('Admin Tools')
      .addItem('Create Initial Triggers', 'setupTriggers')
      .addItem('Delete Triggers', 'removeTriggers')
      .addItem('Validate Setup', 'validateSetup'))
    .addToUi();

  DEBUG && console.log('createMenu: Menu created');
}

// ===== HELPER FUNCTIONS FOR MENU ITEMS =====
// These are called by menu items and need to be in the client script,
// but we provide library versions that can be used by clients.

/**
 * Gets a status summary for the spreadsheet.
 * @param {string} spreadsheetId - The spreadsheet ID
 * @returns {Object} Status summary
 */
function getStatusSummary(spreadsheetId) {
  const ctx = new ExecutionContext(spreadsheetId);
  return ctx.getSummary();
}

/**
 * Validates the spreadsheet configuration.
 * @param {string} spreadsheetId - The spreadsheet ID
 * @returns {Object} Validation results
 */
function validateConfiguration(spreadsheetId) {
  try {
    const ctx = new ExecutionContext(spreadsheetId);
    ctx.validate({ includeFileAccess: true });
    return { valid: true, errors: [] };
  } catch (error) {
    return { valid: false, errors: [error.message] };
  }
}

/**
 * Initializes dropdown data validation for the automation_status column.
 * Sets up dynamic validation rules based on each row's current status.
 * Call this once during initial setup, or to reset validation rules.
 *
 * @param {string} spreadsheetId - The Main Projects File spreadsheet ID
 */
function initializeDropdownValidation(spreadsheetId) {
  console.log('=== initializeDropdownValidation starting ===');

  let ctx;

  try {
    ctx = new ExecutionContext(spreadsheetId);
    ctx.validationService.initializeColumnValidation();
    console.log('=== initializeDropdownValidation completed ===');
  } catch (error) {
    console.error(`initializeDropdownValidation error: ${error.message}`);

    const messageLines = [
      `Error: ${error.message}`,
      '',
      'Function: initializeDropdownValidation',
      `Spreadsheet ID: ${spreadsheetId}`,
      '',
      `Stack: ${error.stack || 'N/A'}`
    ];
    sendAdminErrorNotification(
      spreadsheetId,
      ctx,
      'Initialize Dropdown Validation Failed',
      messageLines.join('\n')
    );

    throw error;
  }
}

/**
 * Refreshes dropdown validation for all project rows.
 * Updates validation rules to match current automation status values.
 *
 * @param {string} spreadsheetId - The Main Projects File spreadsheet ID
 */
function refreshDropdownValidation(spreadsheetId) {
  console.log('=== refreshDropdownValidation starting ===');

  let ctx;

  try {
    ctx = new ExecutionContext(spreadsheetId);
    ctx.validationService.updateAllDropdownValidations();
    console.log('=== refreshDropdownValidation completed ===');
  } catch (error) {
    console.error(`refreshDropdownValidation error: ${error.message}`);

    const messageLines = [
      `Error: ${error.message}`,
      '',
      'Function: refreshDropdownValidation',
      `Spreadsheet ID: ${spreadsheetId}`,
      '',
      `Stack: ${error.stack || 'N/A'}`
    ];
    sendAdminErrorNotification(
      spreadsheetId,
      ctx,
      'Refresh Dropdown Validation Failed',
      messageLines.join('\n')
    );

    throw error;
  }
}

/**
 * Sends an admin error notification for entry-point failures.
 * Tries to use an existing ExecutionContext if available, otherwise builds
 * a minimal NotificationService using only the Config and Directory sheets.
 *
 * This is designed to work even when parts of the data layer (e.g., Codes)
 * are misconfigured and prevent full context initialization.
 *
 * @param {string} spreadsheetId - The Main Projects File spreadsheet ID
 * @param {ExecutionContext|undefined|null} ctx - Existing execution context, if available
 * @param {string} subject - Error subject (without the [Teaming Tool Error] prefix)
 * @param {string} message - Error message/details
 */
function sendAdminErrorNotification(spreadsheetId, ctx, subject, message) {
  try {
    let notificationService = null;

    if (ctx && ctx.notificationService) {
      notificationService = ctx.notificationService;
    } else {
      const sSht = SpreadsheetApp.openById(spreadsheetId);
      const configSheet = sSht.getSheetByName(SHEET_NAMES.CONFIG);
      if (!configSheet) {
        console.error('sendAdminErrorNotification: Config sheet not found, cannot send error email');
        return;
      }

      const directorySheet = sSht.getSheetByName(SHEET_NAMES.DIRECTORY);
      if (!directorySheet) {
        console.error('sendAdminErrorNotification: Directory sheet not found, cannot send error email');
        return;
      }

      const config = new Config(configSheet);
      const directory = new Directory(directorySheet);
      notificationService = new NotificationService(config, directory);
    }

    if (!notificationService) {
      console.error('sendAdminErrorNotification: NotificationService unavailable, skipping error email');
      return;
    }

    notificationService.sendErrorNotification(subject, message);
  } catch (notifyError) {
    console.error(`sendAdminErrorNotification: Failed to send error notification: ${notifyError.message}`);
  }
}

