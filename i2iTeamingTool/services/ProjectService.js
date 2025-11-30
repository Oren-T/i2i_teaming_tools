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

  // ===== VALIDATION METHODS =====

  /**
   * Validates that a project has all required data before processing.
   * Checks required fields and validates directory lookups.
   * @param {Project} project - The project to validate
   * @returns {Object} { valid: boolean, errors: string[] }
   */
  validateProjectData(project) {
    const errors = [];

    // Required field: Project Title
    if (!project.projectName) {
      errors.push('Missing required field: Project Title');
    }

    // Required field: Deadline (Due Date)
    if (!project.dueDate) {
      errors.push('Missing required field: Deadline (Due Date)');
    }

    // Required field: Requested By (must resolve to valid email)
    if (!project.requestedBy) {
      errors.push('Missing required field: Requested By');
    } else {
      const requesterEmail = this.directory.resolveToEmail(project.requestedBy);
      if (!requesterEmail) {
        errors.push(`Requester "${project.requestedBy}" not found in Directory and is not a valid email address`);
      }
    }

    // Required field: Assignee (at least one that resolves to valid email)
    const assignees = project.assignees;
    if (assignees.length === 0) {
      errors.push('Missing required field: Assignee');
    } else {
      const invalidAssignees = [];
      for (const assignee of assignees) {
        const email = this.directory.resolveToEmail(assignee);
        if (!email) {
          invalidAssignees.push(assignee);
        }
      }
      if (invalidAssignees.length > 0) {
        if (invalidAssignees.length === assignees.length) {
          // All assignees are invalid
          errors.push(`No valid assignees found. The following could not be resolved: ${invalidAssignees.join(', ')}`);
        } else {
          // Some assignees are invalid
          errors.push(`Some assignees could not be resolved: ${invalidAssignees.join(', ')}`);
        }
      }
    }

    // Optional field: Category - default to LCAP if missing (handled in processing, not an error)

    return { valid: errors.length === 0, errors };
  }

  /**
   * Builds a detailed error message for notification purposes.
   * @param {Project} project - The project that encountered an error
   * @param {string} errorType - Brief description of the error type
   * @param {string|string[]} errorDetails - The specific error(s)
   * @returns {string} Formatted error message
   */
  buildErrorMessage(project, errorType, errorDetails) {
    const details = Array.isArray(errorDetails) ? errorDetails : [errorDetails];
    const row = project.getRowIndex();

    const lines = [
      `${errorType}`,
      '',
      'Project Details:',
      `  Row: ${row}`,
      `  Title: ${project.projectName || '(not provided)'}`,
      `  Project ID: ${project.projectId || '(not yet assigned)'}`,
      `  Requested By: ${project.requestedBy || '(not provided)'}`,
      `  Assignee(s): ${project.assignee || '(not provided)'}`,
      `  Due Date: ${project.dueDate ? formatDate(project.dueDate) : '(not provided)'}`,
      '',
      'Error Details:'
    ];

    for (const detail of details) {
      lines.push(`  • ${detail}`);
    }

    return lines.join('\n');
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

        // Build detailed error message
        const errorMessage = this.buildErrorMessage(
          project,
          'An unexpected error occurred while processing this project.',
          error.message
        );

        // Send error notification with CC to requester (best-effort)
        const requesterEmail = this.directory.resolveToEmail(project.requestedBy);
        this.notificationService.sendErrorNotification(
          'Project Processing Failed',
          errorMessage,
          { cc: requesterEmail }
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

    // Validate required data FIRST, before any side effects
    const validation = this.validateProjectData(project);
    if (!validation.valid) {
      const errorMessage = this.buildErrorMessage(
        project,
        'Project failed validation. Please correct the following issues and set status back to "Ready" to retry.',
        validation.errors
      );
      console.error(`ProjectService: Validation failed for row ${project.getRowIndex()}: ${validation.errors.join('; ')}`);
      this.setAutomationStatus(project, AUTOMATION_STATUS.ERROR);

      // Send error notification - CC requester if we can resolve their email
      // (they might not be in directory, which is one of the validation errors)
      const requesterEmail = this.directory.resolveToEmail(project.requestedBy);
      this.notificationService.sendErrorNotification(
        'Project Validation Failed',
        errorMessage,
        { cc: requesterEmail }
      );
      return;
    }

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
      // Infer school year from deadline (validated above, so this should always succeed)
      const schoolYear = inferSchoolYear(project.dueDate, this.config.schoolYearStartMonth);
      DEBUG && console.log(`ProjectService: Inferred school year ${schoolYear} from deadline`);

      // Generate new project ID with the school year
      projectId = this.idAllocator.next(schoolYear);
      project.projectId = projectId;
      project.createdAt = new Date();
      project.schoolYear = schoolYear;
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

    // Set default reminder timelines if not already specified
    if (!project.reminderOffsetsRaw) {
      const defaultLabels = this.ctx.codes.getDefaultReminderLabels();
      if (defaultLabels.length > 0) {
        project.reminderOffsets = defaultLabels.join(', ');
        DEBUG && console.log(`ProjectService: Set default reminder labels: ${project.reminderOffsetsRaw}`);
      }
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

        // Build detailed error message
        const errorMessage = this.buildErrorMessage(
          project,
          'An error occurred while processing project updates.',
          error.message
        );

        // Send error notification with CC to requester
        const requesterEmail = this.directory.resolveToEmail(project.requestedBy);
        this.notificationService.sendErrorNotification(
          'Project Update Failed',
          errorMessage,
          { cc: requesterEmail }
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

    // Capture "before" state from calendar event for change detection
    const beforeState = this.getCalendarEventSnapshot(project);

    // Update calendar event
    this.updateCalendarEvent(project);

    // Update project file (Overview tab)
    this.updateProjectFile(project);

    // Re-sync folder sharing (handles assignee changes - idempotent)
    this.shareProjectFolder(project, project.folderId);

    // Set status back to Created
    this.setAutomationStatus(project, AUTOMATION_STATUS.CREATED);

    // Detect what changed and send update notification
    const changes = this.detectProjectChanges(beforeState, project);
    this.notificationService.sendUpdateNotification(project, changes);
  }

  /**
   * Gets a snapshot of the current calendar event state for comparison.
   * @param {Project} project - The project to get calendar state for
   * @returns {Object|null} Snapshot with title, date, guestEmails, or null if unavailable
   */
  getCalendarEventSnapshot(project) {
    const eventId = project.calendarEventId;
    if (!eventId) {
      DEBUG && console.log(`ProjectService: No calendar event ID for ${project.projectId}, cannot snapshot`);
      return null;
    }

    try {
      const calendar = CalendarApp.getDefaultCalendar();
      const event = withBackoff(() => calendar.getEventById(eventId));

      if (!event) {
        DEBUG && console.log(`ProjectService: Calendar event ${eventId} not found`);
        return null;
      }

      // Extract guest emails (normalized to lowercase)
      const guestEmails = event.getGuestList()
        .map(g => g.getEmail().toLowerCase())
        .sort();

      // Get event date - for all-day events, use getAllDayStartDate()
      const eventDate = event.isAllDayEvent()
        ? event.getAllDayStartDate()
        : event.getStartTime();

      return {
        title: event.getTitle(),
        date: eventDate,
        guestEmails: guestEmails
      };
    } catch (error) {
      console.warn(`ProjectService: Could not snapshot calendar event: ${error.message}`);
      return null;
    }
  }

  /**
   * Detects changes between the calendar snapshot and current project state.
   * @param {Object|null} beforeState - The calendar snapshot (or null if unavailable)
   * @param {Project} project - The current project state
   * @returns {Object|null} Changes object or null if comparison not possible
   */
  detectProjectChanges(beforeState, project) {
    // If we couldn't get the before state, we can't detect changes
    if (!beforeState) {
      return null;
    }

    const changes = {
      titleChanged: null,
      dateChanged: null,
      peopleAdded: [],
      peopleRemoved: []
    };

    // Compare title (calendar stores displayTitle format)
    const currentTitle = project.displayTitle;
    if (beforeState.title !== currentTitle) {
      changes.titleChanged = {
        old: beforeState.title,
        new: currentTitle
      };
    }

    // Compare date
    const currentDate = project.dueDate;
    if (beforeState.date && currentDate) {
      if (!isSameDay(beforeState.date, currentDate)) {
        changes.dateChanged = {
          old: beforeState.date,
          new: currentDate
        };
      }
    } else if (beforeState.date !== currentDate) {
      // One is null and the other isn't
      changes.dateChanged = {
        old: beforeState.date,
        new: currentDate
      };
    }

    // Compare people (assignees + requestor)
    const currentEmails = project.getAllRecipientEmails(this.directory)
      .map(e => e.toLowerCase())
      .sort();

    const beforeEmails = new Set(beforeState.guestEmails);
    const currentEmailSet = new Set(currentEmails);

    // Find added people (in current but not in before)
    for (const email of currentEmails) {
      if (!beforeEmails.has(email)) {
        changes.peopleAdded.push(email);
      }
    }

    // Find removed people (in before but not in current)
    for (const email of beforeState.guestEmails) {
      if (!currentEmailSet.has(email)) {
        changes.peopleRemoved.push(email);
      }
    }

    // Check if any changes were detected
    const hasChanges = changes.titleChanged ||
                       changes.dateChanged ||
                       changes.peopleAdded.length > 0 ||
                       changes.peopleRemoved.length > 0;

    return hasChanges ? changes : { noKeyChanges: true };
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

        // Build detailed error message
        const errorMessage = this.buildErrorMessage(
          project,
          'An error occurred while processing the project deletion request.',
          error.message
        );

        // Send error notification with CC to requester
        const requesterEmail = this.directory.resolveToEmail(project.requestedBy);
        this.notificationService.sendErrorNotification(
          'Project Deletion Failed',
          errorMessage,
          { cc: requesterEmail }
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
      const errorMessage = this.buildErrorMessage(
        project,
        'Some issues occurred while sharing the project folder. The project was still created, but manual sharing may be needed.',
        sharingErrors
      );

      // CC the requester so they're aware of sharing issues
      const requesterEmail = this.directory.resolveToEmail(project.requestedBy);
      this.notificationService.sendErrorNotification(
        'Project Folder Sharing Issues',
        errorMessage,
        { cc: requesterEmail }
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

