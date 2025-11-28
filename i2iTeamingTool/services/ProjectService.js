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
        this.setAutomationStatus(project, AUTOMATION_STATUS.ERROR);

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

    // Check if project already has an ID.
    // If so, treat this as a resume/retry operation rather than a new creation.
    let projectId = project.projectId;
    const isResume = project.hasProjectId;
    
    // Track what already existed to determine if this is a "full retry" (everything done)
    const initialState = {
      hasProjectId: project.hasProjectId,
      hasFolderId: !!project.folderId,
      hasFileId: !!project.fileId,
      hasEventId: !!project.calendarEventId
    };

    if (isResume) {
      console.log(`ProjectService: Resuming processing for existing project ${projectId}`);
    } else {
      // Generate new project ID
      projectId = this.idAllocator.next();
      project.projectId = projectId;
      project.createdAt = new Date();
      project.schoolYear = this.config.schoolYear;
    }

    // Check if project has a due date (required for calendar event)
    if (!project.dueDate) {
      const errorMsg = `Project at row ${project.getRowIndex()} is missing a due date. ` +
                       `A deadline is required to create the calendar event.`;
      console.error(`ProjectService: ${errorMsg}`);
      this.setAutomationStatus(project, AUTOMATION_STATUS.ERROR);

      this.notificationService.sendErrorNotification(
        'Project Missing Due Date',
        `${errorMsg}\nProject name: ${project.projectName || '(no name)'}`
      );
      return;
    }

    // Create project folder (idempotent check)
    let folderId = project.folderId;
    if (folderId) {
      DEBUG && console.log(`ProjectService: Folder already exists (${folderId}), skipping creation`);
    } else {
      folderId = this.createProjectFolder(project);
      project.folderId = folderId;
    }

    // Copy template(s) into folder (idempotent check)
    let fileId = project.fileId;
    if (fileId) {
      DEBUG && console.log(`ProjectService: Project file already exists (${fileId}), skipping creation`);
      // Optionally update the existing file to ensure it's in sync
      this.updateProjectFile(project);
    } else {
      this.copyTemplateToFolder(project, folderId);
    }

    // Share folder with assignees and requester (idempotent - addEditor is safe to re-run)
    this.shareProjectFolder(project, folderId);

    // Create calendar event (idempotent check)
    let eventId = project.calendarEventId;
    if (eventId) {
      DEBUG && console.log(`ProjectService: Calendar event already exists (${eventId}), skipping creation`);
      // Optionally update the existing event
      this.updateCalendarEvent(project);
    } else {
      eventId = this.createCalendarEvent(project);
      project.calendarEventId = eventId;
    }

    // Ensure project status is set (for manually added rows)
    if (!project.projectStatus) {
      project.projectStatus = PROJECT_STATUS.PROJECT_ASSIGNED;
    }

    // Update automation status
    this.setAutomationStatus(project, AUTOMATION_STATUS.CREATED);

    DEBUG && console.log(`ProjectService: Created/Resumed project ${projectId} with folder ${folderId} and event ${eventId}`);

    // Determine notification logic
    // If everything already existed, we treat this as a silent success (no email).
    // Otherwise, we assume something was missing/failed previously, so we send the "New Project" email.
    const isFullRetry = initialState.hasProjectId && 
                        initialState.hasFolderId && 
                        initialState.hasFileId && 
                        initialState.hasEventId;

    if (isFullRetry) {
      console.log('ProjectService: All artifacts already existed. Skipping "New Project" email (silent success).');
    } else {
      this.notificationService.sendNewProjectEmail(project);
    }
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
        this.setAutomationStatus(project, AUTOMATION_STATUS.ERROR);

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

    // Update project file (Overview tab)
    this.updateProjectFile(project);

    // Re-sync folder sharing (handles assignee changes - idempotent)
    this.shareProjectFolder(project, project.folderId);

    // Set status back to Created
    this.setAutomationStatus(project, AUTOMATION_STATUS.CREATED);

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
        this.setAutomationStatus(project, AUTOMATION_STATUS.ERROR);

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
    this.setAutomationStatus(project, AUTOMATION_STATUS.DELETED);

    // Hide the row (after flush)
    // Note: We'll hide after flush in the main processing loop
    project._pendingHide = true;
  }

  /**
   * Sets automation status and refreshes validation for the project row.
   * @param {Project} project - The project being updated
   * @param {string} status - The automation status to apply
   */
  setAutomationStatus(project, status) {
    project.automationStatus = status;
    this.updateAutomationValidation(project);
  }

  /**
   * Refreshes automation_status validation for a single project row.
   * @param {Project} project - The project whose row should be updated
   */
  updateAutomationValidation(project) {
    const validationService = this.ctx && this.ctx.validationService;
    if (validationService && typeof validationService.updateDropdownValidation === 'function') {
      validationService.updateDropdownValidation(project);
    }
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
    const fileId = copiedFile.getId();
    
    // Save file ID to project record
    project.fileId = fileId;

    DEBUG && console.log(`ProjectService: Copied template to "${copyName}" (ID: ${fileId})`);

    // Perform token substitution in the copied spreadsheet
    this.updateProjectFile(project);
  }

  /**
   * Shares the project folder with assignees and requester.
   * Uses Advanced Drive Service to suppress Google's default share notification emails.
   * @param {Project} project - The project
   * @param {string} folderId - The folder ID to share
   */
  shareProjectFolder(project, folderId) {
    if (!folderId) {
      DEBUG && console.log('ProjectService: No folder ID provided, skipping sharing');
      return;
    }

    const emailsToShare = new Set();
    const sharingErrors = [];

    // Collect assignee emails
    for (const assignee of project.assignees) {
      const email = this.directory.resolveToEmail(assignee);
      if (email) {
        if (isValidEmail(email)) {
          emailsToShare.add(email.toLowerCase());
        } else {
          const errorMsg = `Invalid email format for assignee "${assignee}": ${email}`;
          console.warn(`ProjectService: ${errorMsg}`);
          sharingErrors.push(errorMsg);
        }
      } else {
        const errorMsg = `Assignee "${assignee}" not found in Directory`;
        console.warn(`ProjectService: ${errorMsg}`);
        sharingErrors.push(errorMsg);
      }
    }

    // Collect requester email
    if (project.requestedBy) {
      const requesterEmail = this.directory.resolveToEmail(project.requestedBy);
      if (requesterEmail) {
        if (isValidEmail(requesterEmail)) {
          emailsToShare.add(requesterEmail.toLowerCase());
        } else {
          const errorMsg = `Invalid email format for requester "${project.requestedBy}": ${requesterEmail}`;
          console.warn(`ProjectService: ${errorMsg}`);
          sharingErrors.push(errorMsg);
        }
      } else {
        const errorMsg = `Requester "${project.requestedBy}" not found in Directory`;
        console.warn(`ProjectService: ${errorMsg}`);
        sharingErrors.push(errorMsg);
      }
    }

    // Share with each person using Advanced Drive Service (suppresses notification emails)
    for (const email of emailsToShare) {
      try {
        withBackoff(() => {
          Drive.Permissions.insert(
            {
              role: 'writer',
              type: 'user',
              value: email
            },
            folderId,
            {
              sendNotificationEmails: false
            }
          );
        });
        DEBUG && console.log(`ProjectService: Shared folder with ${email}`);
      } catch (e) {
        // Check if error is "already has access" - this is not a real error
        if (e.message && e.message.includes('already has access')) {
          DEBUG && console.log(`ProjectService: ${email} already has access to folder`);
        } else {
          const errorMsg = `Could not share folder with ${email}: ${e.message}`;
          console.warn(`ProjectService: ${errorMsg}`);
          sharingErrors.push(errorMsg);
        }
      }
    }

    // Send error notification if there were sharing problems
    if (sharingErrors.length > 0) {
      this.notificationService.sendErrorNotification(
        'Project Folder Sharing Issues',
        `Some issues occurred while sharing folder for project ${project.projectId} (${project.projectName}):\n\n` +
        sharingErrors.map(err => `• ${err}`).join('\n') +
        `\n\nThe project was still created, but manual sharing may be needed.`
      );
    }
  }

  /**
   * Updates the project file's Overview tab with current project details.
   * Can be called during creation or update.
   * @param {Project} project - The project with values to sync
   */
  updateProjectFile(project) {
    const fileId = project.fileId;
    if (!fileId) {
      DEBUG && console.log('ProjectService: No file ID found for project, skipping file update');
      return;
    }

    try {
      // Open the specific project file
      const sSht = withBackoff(() => SpreadsheetApp.openById(fileId));
      this.writeProjectDetailsToSheet(sSht, project);
    } catch (error) {
      console.warn(`ProjectService: Could not update project file ${fileId}: ${error.message}`);
      // Don't throw, just log - we don't want to fail the whole process if the file is missing/locked
    }
  }

  /**
   * Writes project details to the Overview tab of a spreadsheet.
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sSht - The target spreadsheet
   * @param {Project} project - The project data source
   */
  writeProjectDetailsToSheet(sSht, project) {
    const overviewSheet = sSht.getSheetByName('Overview');

    if (!overviewSheet) {
      DEBUG && console.log('ProjectService: No Overview sheet in project file, skipping update');
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

    DEBUG && console.log('ProjectService: Updated project details in Overview tab');
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

    // Set initial project status
    projectData.project_status = PROJECT_STATUS.PROJECT_ASSIGNED;

    // Set automation status to Ready
    projectData.automation_status = AUTOMATION_STATUS.READY;

    // Append the row
    const project = this.projectSheet.appendRow(projectData);

    // Ensure validation matches the initialized Ready status
    this.updateAutomationValidation(project);

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
    // Approach 1: Try namedValues['Email Address'] (Google's standard key when collecting emails)
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

    // Approach 2: Fallback to values[1] (standard position when form collects emails)
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

