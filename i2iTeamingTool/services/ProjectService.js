/**
 * ProjectService class - Handles project lifecycle operations.
 * Creates folders, templates, calendar events, and processes status transitions.
 */
class ProjectService {
  /**
   * Creates a new ProjectService instance.
   * @param {ExecutionContext} ctx - The execution context
   */
  constructor(ctx) {
    this.ctx = ctx;
    this.config = ctx.config;
    this.projectSheet = ctx.projectSheet;
    this.idAllocator = ctx.idAllocator;
    this.directory = ctx.directory;
    this.notificationService = ctx.notificationService;
  }

  // ===== MAIN PROCESSING METHODS =====

  /**
   * Processes all projects with automation_status = 'Ready'.
   * Creates folder, templates, calendar event, sends email, sets status to Created.
   */
  processReadyProjects() {
    const readyProjects = this.projectSheet.getReadyProjects();
    DEBUG && console.log(`ProjectService: Processing ${readyProjects.length} ready project(s)`);

    for (const project of readyProjects) {
      try {
        this.processReadyProject(project);
      } catch (error) {
        console.error(`ProjectService: Error processing project at row ${project.getRowIndex()}: ${error.message}`);
        project.automationStatus = AUTOMATION_STATUS.ERROR;

        // Send error notification
        this.notificationService.sendErrorNotification(
          'Project Processing Failed',
          `Failed to process project at row ${project.getRowIndex()}.\n` +
          `Project: ${project.projectName || '(no name)'}\n` +
          `Error: ${error.message}`
        );
      }
    }
  }

  /**
   * Processes a single ready project.
   * @param {Project} project - The project to process
   */
  processReadyProject(project) {
    DEBUG && console.log(`ProjectService: Processing ready project at row ${project.getRowIndex()}`);

    // Check if project already has an ID (should not happen for Ready status)
    if (project.hasProjectId) {
      const errorMsg = `Project at row ${project.getRowIndex()} already has ID (${project.projectId}) but status is Ready. ` +
                       `This may indicate a duplicate processing attempt or manual error.`;
      console.error(`ProjectService: ${errorMsg}`);
      project.automationStatus = AUTOMATION_STATUS.ERROR;

      // Send error notification to admins
      this.notificationService.sendErrorNotification(
        'Project Already Has ID',
        errorMsg
      );
      return;
    }

    // Check if project has a due date (required for calendar event)
    if (!project.dueDate) {
      const errorMsg = `Project at row ${project.getRowIndex()} is missing a due date. ` +
                       `A deadline is required to create the calendar event.`;
      console.error(`ProjectService: ${errorMsg}`);
      project.automationStatus = AUTOMATION_STATUS.ERROR;

      this.notificationService.sendErrorNotification(
        'Project Missing Due Date',
        `${errorMsg}\nProject name: ${project.projectName || '(no name)'}`
      );
      return;
    }

    // Generate project ID
    const projectId = this.idAllocator.next();
    project.projectId = projectId;

    // Set created timestamp and school year
    project.createdAt = new Date();
    project.schoolYear = this.config.schoolYear;

    // Create project folder
    const folderId = this.createProjectFolder(project);
    project.folderId = folderId;

    // Copy template(s) into folder
    this.copyTemplateToFolder(project, folderId);

    // Create calendar event
    const eventId = this.createCalendarEvent(project);
    project.calendarEventId = eventId;

    // Update automation status
    project.automationStatus = AUTOMATION_STATUS.CREATED;

    DEBUG && console.log(`ProjectService: Created project ${projectId} with folder ${folderId} and event ${eventId}`);

    // Send notification email
    this.notificationService.sendNewProjectEmail(project);
  }

  /**
   * Processes all projects with automation_status = 'Updated'.
   * Re-syncs calendar event and sends update notifications.
   */
  processUpdatedProjects() {
    const updatedProjects = this.projectSheet.getUpdatedProjects();
    DEBUG && console.log(`ProjectService: Processing ${updatedProjects.length} updated project(s)`);

    for (const project of updatedProjects) {
      try {
        this.processUpdatedProject(project);
      } catch (error) {
        console.error(`ProjectService: Error processing update for ${project.projectId}: ${error.message}`);
        project.automationStatus = AUTOMATION_STATUS.ERROR;

        this.notificationService.sendErrorNotification(
          'Project Update Failed',
          `Failed to process update for project ${project.projectId}.\n` +
          `Error: ${error.message}`
        );
      }
    }
  }

  /**
   * Processes a single updated project.
   * @param {Project} project - The project to process
   */
  processUpdatedProject(project) {
    DEBUG && console.log(`ProjectService: Processing updated project ${project.projectId}`);

    // Update calendar event
    this.updateCalendarEvent(project);

    // Set status back to Created
    project.automationStatus = AUTOMATION_STATUS.CREATED;

    // Send update notification
    this.notificationService.sendUpdateNotification(project);
  }

  /**
   * Processes all projects with automation_status = 'Delete (Notify)' or 'Delete (Don't Notify)'.
   * Cancels calendar event, optionally notifies, and hides the row.
   */
  processDeleteRequests() {
    const deleteProjects = this.projectSheet.getPendingDeleteProjects();
    DEBUG && console.log(`ProjectService: Processing ${deleteProjects.length} delete request(s)`);

    for (const project of deleteProjects) {
      try {
        this.processDeleteRequest(project);
      } catch (error) {
        console.error(`ProjectService: Error processing delete for ${project.projectId}: ${error.message}`);
        project.automationStatus = AUTOMATION_STATUS.ERROR;

        this.notificationService.sendErrorNotification(
          'Project Deletion Failed',
          `Failed to process deletion for project ${project.projectId}.\n` +
          `Error: ${error.message}`
        );
      }
    }
  }

  /**
   * Processes a single delete request.
   * @param {Project} project - The project to delete
   */
  processDeleteRequest(project) {
    DEBUG && console.log(`ProjectService: Processing delete request for ${project.projectId}`);

    const shouldNotify = project.shouldNotifyOnDelete;

    // Cancel calendar event
    this.cancelCalendarEvent(project);

    // Send cancellation notification if requested
    if (shouldNotify) {
      this.notificationService.sendCancellationNotification(project);
    }

    // Update status
    project.automationStatus = AUTOMATION_STATUS.DELETED;

    // Hide the row (after flush)
    // Note: We'll hide after flush in the main processing loop
    project._pendingHide = true;
  }

  // ===== FOLDER & TEMPLATE METHODS =====

  /**
   * Creates a project folder in the parent folder.
   * @param {Project} project - The project
   * @returns {string} The created folder ID
   */
  createProjectFolder(project) {
    const parentFolderId = this.config.parentFolderId;
    if (!parentFolderId) {
      throw new Error('Parent Folder ID not configured');
    }

    const parentFolder = withBackoff(() => DriveApp.getFolderById(parentFolderId));
    const folderName = `${project.projectName} [${project.projectId}]`;

    const newFolder = withBackoff(() => parentFolder.createFolder(folderName));
    const folderId = newFolder.getId();

    DEBUG && console.log(`ProjectService: Created folder "${folderName}" with ID ${folderId}`);

    return folderId;
  }

  /**
   * Copies the project template into the project folder.
   * Performs token substitution on the template.
   * @param {Project} project - The project
   * @param {string} folderId - The destination folder ID
   */
  copyTemplateToFolder(project, folderId) {
    const templateId = this.config.projectTemplateId;
    if (!templateId) {
      DEBUG && console.log('ProjectService: No project template configured, skipping');
      return;
    }

    const folder = withBackoff(() => DriveApp.getFolderById(folderId));
    const templateFile = withBackoff(() => DriveApp.getFileById(templateId));

    // Create copy with project name
    const copyName = `${project.projectName} - Project File`;
    const copiedFile = withBackoff(() => templateFile.makeCopy(copyName, folder));

    DEBUG && console.log(`ProjectService: Copied template to "${copyName}"`);

    // Perform token substitution in the copied spreadsheet
    this.substituteTemplateTokens(copiedFile.getId(), project);
  }

  /**
   * Substitutes tokens in the template file's Overview tab.
   * @param {string} fileId - The copied template file ID
   * @param {Project} project - The project with values to substitute
   */
  substituteTemplateTokens(fileId, project) {
    try {
      const ss = withBackoff(() => SpreadsheetApp.openById(fileId));
      const overviewSheet = ss.getSheetByName('Overview');

      if (!overviewSheet) {
        DEBUG && console.log('ProjectService: No Overview sheet in template, skipping substitution');
        return;
      }

      // Map of row labels to project values
      const valueMap = {
        'School Year': project.schoolYear,
        'Goal #': project.goalNumber,
        'Action #': project.actionNumber,
        'Category (default is LCAP)': project.category,
        'Title': project.projectName,
        'Description': project.description,
        'Assigned to': project.assignee,
        'Requested by': project.requestedBy,
        'Deadline': project.dueDate ? formatDate(project.dueDate) : ''
      };

      // Read column A to find labels
      const labelRange = overviewSheet.getRange('A2:A10');
      const labels = labelRange.getValues();

      for (let i = 0; i < labels.length; i++) {
        const label = String(labels[i][0]).trim();
        if (valueMap[label] !== undefined) {
          const row = i + 2; // Row 2 is index 0
          overviewSheet.getRange(row, 2).setValue(valueMap[label]);
        }
      }

      DEBUG && console.log('ProjectService: Substituted tokens in template Overview tab');

    } catch (error) {
      console.warn(`ProjectService: Could not substitute template tokens: ${error.message}`);
    }
  }

  // ===== CALENDAR METHODS =====

  /**
   * Creates a calendar event for the project.
   * Uses the bot account's default calendar.
   * @param {Project} project - The project
   * @returns {string} The created event ID
   */
  createCalendarEvent(project) {
    const calendar = CalendarApp.getDefaultCalendar();

    const title = project.displayTitle;
    const dueDate = project.dueDate;

    if (!dueDate) {
      console.warn(`ProjectService: No due date for project ${project.projectId}, skipping calendar event`);
      return '';
    }

    const description = this.buildCalendarDescription(project);
    const guestEmails = project.getAllRecipientEmails(this.directory);

    // Create event with guests and send invites at creation time
    const event = withBackoff(() =>
      calendar.createAllDayEvent(title, dueDate, {
        description,
        guests: guestEmails.join(','),
        sendInvites: true
      })
    );

    const eventId = event.getId();
    DEBUG && console.log(`ProjectService: Created calendar event ${eventId} for ${project.projectId} with ${guestEmails.length} guest(s)`);

    return eventId;
  }

  /**
   * Updates an existing calendar event.
   * @param {Project} project - The project with updated details
   */
  updateCalendarEvent(project) {
    const eventId = project.calendarEventId;
    if (!eventId) {
      DEBUG && console.log(`ProjectService: No calendar event ID for ${project.projectId}`);
      return;
    }

    const calendar = CalendarApp.getDefaultCalendar();

    try {
      const event = withBackoff(() => calendar.getEventById(eventId));
      if (!event) {
        console.warn(`ProjectService: Calendar event ${eventId} not found`);
        return;
      }

      // Update title
      event.setTitle(project.displayTitle);

      // Update date if changed
      if (project.dueDate) {
        event.setAllDayDate(project.dueDate);
      }

      // Update description
      event.setDescription(this.buildCalendarDescription(project));

      // Update guests - remove old, add new
      const currentGuests = event.getGuestList().map(g => g.getEmail().toLowerCase());
      const newGuests = project.getAllRecipientEmails(this.directory);

      // Remove guests not in new list
      for (const guest of currentGuests) {
        if (!newGuests.includes(guest)) {
          event.removeGuest(guest);
        }
      }

      // Add new guests not in current list
      for (const guest of newGuests) {
        if (!currentGuests.includes(guest.toLowerCase())) {
          try {
            event.addGuest(guest);
          } catch (e) {
            console.warn(`ProjectService: Could not add guest ${guest}: ${e.message}`);
          }
        }
      }

      DEBUG && console.log(`ProjectService: Updated calendar event ${eventId}`);

    } catch (error) {
      console.error(`ProjectService: Error updating calendar event: ${error.message}`);
      throw error;
    }
  }

  /**
   * Cancels a calendar event.
   * @param {Project} project - The project
   */
  cancelCalendarEvent(project) {
    const eventId = project.calendarEventId;
    if (!eventId) {
      DEBUG && console.log(`ProjectService: No calendar event ID for ${project.projectId}`);
      return;
    }

    const calendar = CalendarApp.getDefaultCalendar();

    try {
      const event = withBackoff(() => calendar.getEventById(eventId));
      if (event) {
        event.deleteEvent();
        DEBUG && console.log(`ProjectService: Cancelled calendar event ${eventId}`);
      }
    } catch (error) {
      console.warn(`ProjectService: Error cancelling calendar event: ${error.message}`);
    }
  }

  /**
   * Builds the calendar event description.
   * @param {Project} project - The project
   * @returns {string} Event description
   */
  buildCalendarDescription(project) {
    const lines = [
      `Project: ${project.projectName}`,
      `Project ID: ${project.projectId}`,
      `Category: ${project.category}`,
      `Requested by: ${project.requestedBy}`,
      `Assigned to: ${project.assignee}`,
      '',
      `Description: ${project.description}`,
      '',
      `Project Folder: ${project.folderUrl}`
    ];

    return lines.join('\n');
  }

  // ===== FORM RESPONSE PROCESSING =====

  /**
   * Normalizes a form submission and appends to the Projects sheet.
   * @param {Object} event - The form submission event
   * @returns {Project} The created project
   */
  normalizeAndAppendFormResponse(event) {
    DEBUG && console.log('ProjectService: Processing form submission');

    const namedValues = event.namedValues || {};
    const values = event.values || [];

    // Map form fields to project columns
    const projectData = {};

    for (const [formField, internalKey] of Object.entries(FORM_FIELD_MAP)) {
      const fieldValues = namedValues[formField];
      if (fieldValues && fieldValues.length > 0) {
        // Handle multi-select (comma-join) vs single value
        projectData[internalKey] = fieldValues.join(', ');
      }
    }

    // Set requested_by from form submitter's email
    // Try multiple approaches to find the submitter email
    const submitterEmail = this.extractSubmitterEmail(namedValues, values);
    if (submitterEmail) {
      // Look up the name from directory, fall back to email
      const submitterName = this.directory.getNameByEmail(submitterEmail);
      projectData.requested_by = submitterName || submitterEmail;
      DEBUG && console.log(`ProjectService: Form submitted by ${projectData.requested_by}`);
    } else {
      console.warn('ProjectService: Could not determine form submitter email');
    }

    // Set default reminder offsets from Codes sheet
    const defaultOffsets = this.ctx.codes.getDefaultReminderOffsets();
    if (defaultOffsets.length > 0) {
      projectData.reminder_offsets = defaultOffsets.join(',');
      DEBUG && console.log(`ProjectService: Set default reminder offsets: ${projectData.reminder_offsets}`);
    }

    // Set automation status to Ready
    projectData.automation_status = AUTOMATION_STATUS.READY;

    // Append the row
    const project = this.projectSheet.appendRow(projectData);

    DEBUG && console.log(`ProjectService: Appended form response as row ${project.getRowIndex()}`);

    return project;
  }

  /**
   * Extracts the form submitter's email from the event data.
   * Tries multiple approaches: namedValues['Email Address'], then values[1].
   * @param {Object} namedValues - The namedValues from the form event
   * @param {Array} values - The values array from the form event
   * @returns {string|null} The submitter email or null if not found
   */
  extractSubmitterEmail(namedValues, values) {
    // Approach B: Try namedValues['Email Address'] (Google's standard key when collecting emails)
    const emailKeys = ['Email Address', 'Email address', 'email address', 'Email'];
    for (const key of emailKeys) {
      if (namedValues[key] && namedValues[key].length > 0) {
        const email = namedValues[key][0].trim();
        if (email.includes('@')) {
          DEBUG && console.log(`ProjectService: Found submitter email via namedValues['${key}']`);
          return email;
        }
      }
    }

    // Approach A: Fallback to values[1] (standard position when form collects emails)
    // values[0] = timestamp, values[1] = email (when "Collect email addresses" is enabled)
    if (values.length > 1 && values[1]) {
      const email = String(values[1]).trim();
      if (email.includes('@')) {
        DEBUG && console.log('ProjectService: Found submitter email via values[1]');
        return email;
      }
    }

    // Neither approach found a valid email
    console.warn('ProjectService: No valid email found in namedValues or values[1]. ' +
                 'Ensure the form is configured to collect email addresses.');
    return null;
  }

  /**
   * Hides rows for deleted projects (called after flush).
   */
  hideDeletedRows() {
    const projects = this.projectSheet.getProjects();

    for (const project of projects) {
      if (project._pendingHide) {
        this.projectSheet.hideRow(project);
        delete project._pendingHide;
      }
    }
  }
}

