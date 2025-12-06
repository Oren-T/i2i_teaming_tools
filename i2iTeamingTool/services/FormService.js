/**
 * FormService class - Handles Google Form dropdown synchronization.
 * Syncs form dropdown options with Directory and Codes sheets.
 */
class FormService {
  /**
   * Creates a new FormService instance.
   * @param {ExecutionContext} ctx - The execution context
   */
  constructor(ctx) {
    this.ctx = ctx;
    this.config = ctx.config;
    this.directory = ctx.directory;
    this.codes = ctx.codes;
    this.notificationService = ctx.notificationService;
    this.form = null;
  }

  /**
   * Gets the Google Form instance (lazy loaded).
   * @returns {GoogleAppsScript.Forms.Form|null} The form or null
   */
  getForm() {
    if (this.form) {
      return this.form;
    }

    const formId = this.config.formId;
    if (!formId) {
      DEBUG && console.log('FormService: No Form ID configured');
      return null;
    }

    try {
      this.form = withBackoff(() => FormApp.openById(formId));
      return this.form;
    } catch (error) {
      console.error(`FormService: Could not open form: ${error.message}`);
      try {
        const lines = [
          'The Google Form used by the Teaming Tool could not be opened.',
          '',
          `Form ID: ${formId}`,
          '',
          `Error: ${error.message}`,
          `Stack: ${error.stack || 'N/A'}`
        ];

        this.notificationService.sendErrorNotification(
          'Form Access Failed',
          lines.join('\n')
        );
      } catch (notifyError) {
        console.error(`FormService: Failed to send form access error notification: ${notifyError.message}`);
      }
      return null;
    }
  }

  /**
   * Syncs all dropdowns in the form.
   */
  syncAllDropdowns() {
    console.log('FormService: Syncing all dropdowns');

    const form = this.getForm();
    if (!form) {
      console.warn('FormService: No form available, skipping sync');
      return;
    }

    this.syncAssigneeDropdown();
    this.syncCategoryDropdown();

    console.log('FormService: Dropdown sync complete');
  }

  /**
   * Syncs the "Assigned to" dropdown with active staff from Directory.
   */
  syncAssigneeDropdown() {
    const form = this.getForm();
    if (!form) return;

    const staffNames = this.directory.getActiveStaffNames();
    if (staffNames.length === 0) {
      console.warn('FormService: No active staff found in Directory');
      return;
    }

    // Find the "Assigned to" question
    const item = this.findFormItem(form, 'Assigned to');
    if (!item) {
      DEBUG && console.log('FormService: "Assigned to" question not found in form');
      return;
    }

    // Update choices based on item type
    try {
      if (item.getType() === FormApp.ItemType.CHECKBOX) {
        const checkboxItem = item.asCheckboxItem();
        const choices = staffNames.map(name => checkboxItem.createChoice(name));
        checkboxItem.setChoices(choices);
        DEBUG && console.log(`FormService: Updated "Assigned to" checkbox with ${staffNames.length} options`);

      } else if (item.getType() === FormApp.ItemType.LIST) {
        const listItem = item.asListItem();
        const choices = staffNames.map(name => listItem.createChoice(name));
        listItem.setChoices(choices);
        DEBUG && console.log(`FormService: Updated "Assigned to" list with ${staffNames.length} options`);

      } else if (item.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
        const mcItem = item.asMultipleChoiceItem();
        const choices = staffNames.map(name => mcItem.createChoice(name));
        mcItem.setChoices(choices);
        DEBUG && console.log(`FormService: Updated "Assigned to" multiple choice with ${staffNames.length} options`);

      } else {
        console.warn(`FormService: "Assigned to" question is not a supported type: ${item.getType()}`);
      }
    } catch (error) {
      console.error(`FormService: Error updating "Assigned to": ${error.message}`);
      try {
        const lines = [
          'An error occurred while syncing the "Assigned to" dropdown in the Google Form.',
          '',
          `Form ID: ${this.config.formId || '(not configured)'}`,
          `Question title: ${item ? item.getTitle() : 'Assigned to'}`,
          '',
          `Error: ${error.message}`,
          `Stack: ${error.stack || 'N/A'}`
        ];

        this.notificationService.sendErrorNotification(
          'Form Dropdown Sync Failed (Assigned to)',
          lines.join('\n')
        );
      } catch (notifyError) {
        console.error(`FormService: Failed to send "Assigned to" sync error notification: ${notifyError.message}`);
      }
    }
  }

  /**
   * Syncs the "Category" dropdown with categories from Codes sheet.
   */
  syncCategoryDropdown() {
    const form = this.getForm();
    if (!form) return;

    const categories = this.codes.getCategories();
    if (categories.length === 0) {
      console.warn('FormService: No categories found in Codes sheet');
      return;
    }

    // Find the "Category" question
    const item = this.findFormItem(form, 'Category');
    if (!item) {
      DEBUG && console.log('FormService: "Category" question not found in form');
      return;
    }

    try {
      if (item.getType() === FormApp.ItemType.LIST) {
        const listItem = item.asListItem();
        const choices = categories.map(cat => listItem.createChoice(cat));
        listItem.setChoices(choices);
        DEBUG && console.log(`FormService: Updated "Category" list with ${categories.length} options`);

      } else if (item.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
        const mcItem = item.asMultipleChoiceItem();
        const choices = categories.map(cat => mcItem.createChoice(cat));
        mcItem.setChoices(choices);
        DEBUG && console.log(`FormService: Updated "Category" multiple choice with ${categories.length} options`);

      } else if (item.getType() === FormApp.ItemType.CHECKBOX) {
        const checkboxItem = item.asCheckboxItem();
        const choices = categories.map(cat => checkboxItem.createChoice(cat));
        checkboxItem.setChoices(choices);
        DEBUG && console.log(`FormService: Updated "Category" checkbox with ${categories.length} options`);

      } else {
        console.warn(`FormService: "Category" question is not a supported type: ${item.getType()}`);
      }
    } catch (error) {
      console.error(`FormService: Error updating "Category": ${error.message}`);
      try {
        const lines = [
          'An error occurred while syncing the "Category" dropdown in the Google Form.',
          '',
          `Form ID: ${this.config.formId || '(not configured)'}`,
          `Question title: ${item ? item.getTitle() : 'Category'}`,
          '',
          `Error: ${error.message}`,
          `Stack: ${error.stack || 'N/A'}`
        ];

        this.notificationService.sendErrorNotification(
          'Form Dropdown Sync Failed (Category)',
          lines.join('\n')
        );
      } catch (notifyError) {
        console.error(`FormService: Failed to send "Category" sync error notification: ${notifyError.message}`);
      }
    }
  }

  /**
   * Finds a form item by title (case-insensitive).
   * @param {GoogleAppsScript.Forms.Form} form - The form
   * @param {string} title - The item title to find
   * @returns {GoogleAppsScript.Forms.Item|null} The item or null
   */
  findFormItem(form, title) {
    const items = form.getItems();
    const searchTitle = title.toLowerCase();

    for (const item of items) {
      if (item.getTitle().toLowerCase() === searchTitle) {
        return item;
      }
    }

    return null;
  }

  /**
   * Gets a summary of current form structure.
   * Useful for debugging.
   * @returns {Object} Form structure info
   */
  getFormSummary() {
    const form = this.getForm();
    if (!form) {
      return { error: 'No form available' };
    }

    const items = form.getItems();
    const summary = {
      title: form.getTitle(),
      itemCount: items.length,
      items: items.map(item => ({
        title: item.getTitle(),
        type: item.getType().toString(),
        id: item.getId()
      }))
    };

    return summary;
  }

  /**
   * Validates that the form has expected questions.
   * @returns {Object} Validation results
   */
  validateFormStructure() {
    const form = this.getForm();
    if (!form) {
      return { valid: false, errors: ['Form not available'] };
    }

    const expectedFields = Object.keys(FORM_FIELD_MAP);
    const errors = [];

    for (const fieldName of expectedFields) {
      const item = this.findFormItem(form, fieldName);
      if (!item) {
        errors.push(`Missing form question: "${fieldName}"`);
      }
    }

    return {
      valid: errors.length === 0,
      errors
    };
  }
}

