/**
 * Client Script for the i2i Teaming Tool.
 *
 * This is a thin client script that delegates all business logic to the
 * i2iTT library. Each district gets their own copy of this script,
 * configured with their specific spreadsheet ID.
 *
 * SETUP INSTRUCTIONS:
 * 1. Create this script as a bound script to the Main Projects File spreadsheet
 * 2. Add the i2iTT library (Script ID: [YOUR_LIBRARY_SCRIPT_ID])
 * 3. Update SPREADSHEET_ID below with the spreadsheet's ID
 * 4. Run setupTriggers() once to create the time-driven triggers
 */

// ===== CONFIGURATION =====
// Replace with this district's Main Projects File spreadsheet ID
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// ===== TRIGGER HANDLERS =====

/**
 * Handles the onOpen event. Creates the custom menu.
 * @param {Object} e - The event object
 */
function onOpen(e) {
  i2iTT.createMenu();
}

/**
 * Handles the onEdit event.
 * @param {Object} e - The event object
 */
function onEdit(e) {
  i2iTT.handleEdit(SPREADSHEET_ID, e);
}

/**
 * Handles form submission events.
 * This is called by an installable trigger on the spreadsheet.
 * @param {Object} e - The form submission event object
 */
function onFormSubmit(e) {
  i2iTT.handleFormSubmission(SPREADSHEET_ID, e);
}

/**
 * Time-driven trigger handler for the 10-minute batch job.
 * Processes Ready, Updated, and Delete requests.
 */
function onBatchTrigger() {
  i2iTT.processNewProjects(SPREADSHEET_ID);
}

/**
 * Time-driven trigger handler for the daily 8am job.
 * Runs maintenance tasks: reminders, status digest, late marking, calendar sync.
 */
function onDailyTrigger() {
  i2iTT.runDailyMaintenance(SPREADSHEET_ID);
}

// ===== MENU HANDLERS =====

/**
 * Manual trigger for processing projects immediately.
 * Called from the Teaming Tool menu.
 */
function manualRunNow() {
  const ui = SpreadsheetApp.getUi();

  try {
    i2iTT.processNewProjects(SPREADSHEET_ID);
    ui.alert('Success', 'Projects processed successfully.', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', `Failed to process projects: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Manual trigger for syncing form dropdowns.
 * Called from the Teaming Tool menu.
 */
function manualSyncDropdowns() {
  const ui = SpreadsheetApp.getUi();

  try {
    i2iTT.syncFormDropdowns(SPREADSHEET_ID);
    ui.alert('Success', 'Form dropdowns synced successfully.', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', `Failed to sync dropdowns: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Manual trigger for refreshing permissions.
 * Called from the Teaming Tool menu.
 */
function manualRefreshPermissions() {
  const ui = SpreadsheetApp.getUi();

  try {
    i2iTT.refreshPermissions(SPREADSHEET_ID);
    ui.alert('Success', 'Permissions refreshed successfully.', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', `Failed to refresh permissions: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Shows a status summary dialog.
 * Called from the Teaming Tool menu.
 */
function showStatusSummary() {
  const ui = SpreadsheetApp.getUi();

  try {
    const summary = i2iTT.getStatusSummary(SPREADSHEET_ID);

    const message = [
      `District: ${summary.config.districtId}`,
      `School Year: ${summary.config.schoolYear}`,
      `Next Serial: ${summary.config.nextSerial}`,
      '',
      `Total Projects: ${summary.projectCount}`,
      `Ready for Processing: ${summary.readyCount}`,
      `Active (Created): ${summary.createdCount}`,
      '',
      `Staff in Directory: ${summary.staffCount}`,
      '',
      `Last Updated: ${summary.timestamp}`
    ].join('\n');

    ui.alert('Status Summary', message, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', `Failed to get status: ${error.message}`, ui.ButtonSet.OK);
  }
}

// ===== SETUP FUNCTIONS =====

/**
 * Sets up all required triggers for the Teaming Tool.
 * Run this function once after initial setup.
 */
function setupTriggers() {
  const ui = SpreadsheetApp.getUi();

  // Remove existing triggers first
  const existingTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of existingTriggers) {
    const handlerFunction = trigger.getHandlerFunction();
    if (['onBatchTrigger', 'onDailyTrigger', 'onFormSubmit'].includes(handlerFunction)) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Create 10-minute batch trigger
  ScriptApp.newTrigger('onBatchTrigger')
    .timeBased()
    .everyMinutes(10)
    .create();

  // Create daily 8am trigger
  ScriptApp.newTrigger('onDailyTrigger')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  // Create form submit trigger (installable trigger on the spreadsheet)
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SPREADSHEET_ID)
    .onFormSubmit()
    .create();

  ui.alert('Success', 'Triggers have been set up:\n\n' +
    '• 10-minute batch processing\n' +
    '• Daily 8am maintenance\n' +
    '• Form submission handler', ui.ButtonSet.OK);
}

/**
 * Removes all project triggers.
 * Use this to clean up before re-running setupTriggers.
 */
function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;

  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
    count++;
  }

  const ui = SpreadsheetApp.getUi();
  ui.alert('Triggers Removed', `Removed ${count} trigger(s).`, ui.ButtonSet.OK);
}

/**
 * Validates the spreadsheet configuration.
 * Use this to check if everything is set up correctly.
 */
function validateSetup() {
  const ui = SpreadsheetApp.getUi();

  try {
    const result = i2iTT.validateConfiguration(SPREADSHEET_ID);

    if (result.valid) {
      ui.alert('Validation Passed', 'All configuration is valid!', ui.ButtonSet.OK);
    } else {
      ui.alert('Validation Failed',
        'Configuration errors:\n\n' + result.errors.join('\n'),
        ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('Error', `Validation failed: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Initializes dropdown validation for the Automation Status column.
 * Sets up dynamic rules that change based on each row's current status.
 */
function initializeDropdowns() {
  const ui = SpreadsheetApp.getUi();

  try {
    i2iTT.initializeDropdownValidation(SPREADSHEET_ID);
    ui.alert('Success', 'Dropdown validation initialized successfully.', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', `Failed to initialize dropdowns: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Refreshes dropdown validation for all rows.
 * Use this if dropdown options seem out of sync.
 */
function refreshDropdowns() {
  const ui = SpreadsheetApp.getUi();

  try {
    i2iTT.refreshDropdownValidation(SPREADSHEET_ID);
    ui.alert('Success', 'Dropdown validation refreshed successfully.', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', `Failed to refresh dropdowns: ${error.message}`, ui.ButtonSet.OK);
  }
}

