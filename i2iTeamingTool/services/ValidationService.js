/**
 * ValidationService class - Manages dropdown data validation rules.
 * Dynamically updates automation_status dropdown options based on current state.
 */
class ValidationService {
  /**
   * Creates a new ValidationService instance.
   * @param {ExecutionContext} ctx - The execution context
   */
  constructor(ctx) {
    this.ctx = ctx;
    this.projectSheet = ctx.projectSheet;
  }

  /**
   * Updates data validation for all project rows based on their current automation status.
   * Should be called after processing projects to ensure dropdown options reflect current state.
   */
  updateAllDropdownValidations() {
    DEBUG && console.log('ValidationService: Updating dropdown validations for all rows');

    const sheet = this.projectSheet.getSheet();
    const projects = this.projectSheet.getProjects();
    const statusColIndex = this.projectSheet.getColumnIndex('automation_status');

    if (statusColIndex === undefined) {
      console.warn('ValidationService: automation_status column not found');
      return;
    }

    // Column is 1-based in Sheets API
    const statusCol = statusColIndex + 1;
    let updated = 0;

    for (const project of projects) {
      const row = project.getRowIndex();
      const currentStatus = project.automationStatus;
      const allowedValues = this.getAllowedStatusValues(currentStatus);

      this.setDropdownValidation(sheet, row, statusCol, allowedValues);
      updated++;
    }

    // Also handle any blank rows at the end (rows without projects)
    const lastRow = sheet.getLastRow();
    const dataRowCount = projects.length + 2; // +2 for header rows

    for (let row = dataRowCount + 1; row <= lastRow; row++) {
      // Check if this row has any content
      const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      const hasContent = rowData.some(cell => cell !== '' && cell !== null);

      if (!hasContent) {
        // Completely blank row - allow only Ready
        this.setDropdownValidation(sheet, row, statusCol, [AUTOMATION_STATUS.READY]);
      }
    }

    DEBUG && console.log(`ValidationService: Updated ${updated} row validations`);
  }

  /**
   * Updates data validation for a single project row.
   * @param {Project} project - The project to update
   */
  updateDropdownValidation(project) {
    const sheet = this.projectSheet.getSheet();
    const statusColIndex = this.projectSheet.getColumnIndex('automation_status');

    if (statusColIndex === undefined) {
      return;
    }

    const row = project.getRowIndex();
    const statusCol = statusColIndex + 1;
    const allowedValues = this.getAllowedStatusValues(project.automationStatus);

    this.setDropdownValidation(sheet, row, statusCol, allowedValues);
  }

  /**
   * Gets the allowed automation status values based on current status.
   *
   * Rules:
   * - Blank row: can only select "Ready"
   * - Created: can select "Created", "Updated", "Delete (Notify)", "Delete (Don't Notify)"
   * - Error: can select "Ready" (to retry)
   * - Other states (Ready, Updated, Delete *, Deleted): Locked (current value only)
   *
   * @param {string} currentStatus - The current automation status
   * @returns {string[]} Array of allowed status values
   */
  getAllowedStatusValues(currentStatus) {
    switch (currentStatus) {
      case AUTOMATION_STATUS.BLANK:
      case '':
        // Blank row - user can only set to Ready
        return [AUTOMATION_STATUS.READY];

      case AUTOMATION_STATUS.CREATED:
        // Created - user can update, request delete, or keep as is
        return [
          AUTOMATION_STATUS.CREATED,
          AUTOMATION_STATUS.UPDATED,
          AUTOMATION_STATUS.DELETE_NOTIFY,
          AUTOMATION_STATUS.DELETE_NO_NOTIFY
        ];

      case AUTOMATION_STATUS.ERROR:
        // Error - user can retry by setting back to Ready
        return [AUTOMATION_STATUS.ERROR, AUTOMATION_STATUS.READY];

      case AUTOMATION_STATUS.READY:
        // Ready - locked, automation will process it
        return [AUTOMATION_STATUS.READY];

      case AUTOMATION_STATUS.UPDATED:
        // Updated - locked, automation will process it
        return [AUTOMATION_STATUS.UPDATED];

      case AUTOMATION_STATUS.DELETE_NOTIFY:
      case AUTOMATION_STATUS.DELETE_NO_NOTIFY:
        // Delete pending - locked, automation will process it
        return [currentStatus];

      case AUTOMATION_STATUS.DELETED:
        // Deleted - terminal state, no changes allowed
        return [AUTOMATION_STATUS.DELETED];

      default:
        // Unknown status - allow current value only
        return [currentStatus || AUTOMATION_STATUS.READY];
    }
  }

  /**
   * Sets dropdown data validation on a specific cell.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet
   * @param {number} row - Row number (1-based)
   * @param {number} col - Column number (1-based)
   * @param {string[]} allowedValues - Array of allowed dropdown values
   */
  setDropdownValidation(sheet, row, col, allowedValues) {
    const cell = sheet.getRange(row, col);

    if (allowedValues.length === 0) {
      // No validation - clear any existing
      cell.clearDataValidations();
      return;
    }

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(allowedValues, true) // true = show dropdown
      .setAllowInvalid(false) // Reject invalid values
      .setHelpText(this.getHelpText(allowedValues))
      .build();

    cell.setDataValidation(rule);
  }

  /**
   * Generates help text for the dropdown based on allowed values.
   * @param {string[]} allowedValues - The allowed values
   * @returns {string} Help text
   */
  getHelpText(allowedValues) {
    if (allowedValues.length === 1) {
      if (allowedValues[0] === AUTOMATION_STATUS.READY) {
        return 'Set to "Ready" when the row is complete and ready for processing.';
      }
      return 'This field is locked while automation is processing.';
    }

    if (allowedValues.includes(AUTOMATION_STATUS.UPDATED)) {
      return 'Choose "Updated" to re-sync calendar event, or "Delete" to cancel the project.';
    }

    return 'Select a valid status from the dropdown.';
  }

  /**
   * Sets up initial data validation for the automation_status column.
   * Creates a default validation rule for all data rows.
   */
  initializeColumnValidation() {
    DEBUG && console.log('ValidationService: Initializing column validation');

    const sheet = this.projectSheet.getSheet();
    const statusColIndex = this.projectSheet.getColumnIndex('automation_status');

    if (statusColIndex === undefined) {
      console.warn('ValidationService: automation_status column not found');
      return;
    }

    const statusCol = statusColIndex + 1;
    const lastRow = Math.max(sheet.getLastRow(), 3); // At least row 3

    // Default validation for new/blank rows
    const defaultValues = [AUTOMATION_STATUS.READY];
    const defaultRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(defaultValues, true)
      .setAllowInvalid(false)
      .setHelpText('Set to "Ready" when the row is complete and ready for processing.')
      .build();

    // Apply to all data rows (starting at row 3)
    const dataRange = sheet.getRange(3, statusCol, lastRow - 2, 1);
    dataRange.setDataValidation(defaultRule);

    // Then update individual rows based on their current status
    this.updateAllDropdownValidations();

    DEBUG && console.log('ValidationService: Column validation initialized');
  }

  /**
   * Clears all data validation from the automation_status column.
   * Use with caution - mainly for debugging/reset.
   */
  clearColumnValidation() {
    const sheet = this.projectSheet.getSheet();
    const statusColIndex = this.projectSheet.getColumnIndex('automation_status');

    if (statusColIndex === undefined) {
      return;
    }

    const statusCol = statusColIndex + 1;
    const lastRow = sheet.getLastRow();

    if (lastRow > 2) {
      const dataRange = sheet.getRange(3, statusCol, lastRow - 2, 1);
      dataRange.clearDataValidations();
    }

    DEBUG && console.log('ValidationService: Column validation cleared');
  }
}

