/**
 * MaintenanceService class - Handles daily maintenance tasks.
 * Sends reminders, detects status changes, marks late projects, and syncs calendar events.
 */
class MaintenanceService {
  /**
   * Creates a new MaintenanceService instance.
   * @param {ExecutionContext} ctx - The execution context
   */
  constructor(ctx) {
    this.ctx = ctx;
    this.config = ctx.config;
    this.projectSheet = ctx.projectSheet;
    this.snapshotSheet = ctx.snapshotSheet;
    this.directory = ctx.directory;
    this.codes = ctx.codes;
    this.notificationService = ctx.notificationService;
    this.today = ctx.now;
  }

  /**
   * Runs all daily maintenance tasks.
   */
  runDailyMaintenance() {
    console.log('MaintenanceService: Starting daily maintenance');

    try {
      // 1. Send reminders for upcoming deadlines
      this.sendReminders();

      // 2. Mark projects as late if due date is today and not completed
      // (Must happen BEFORE status change detection so Late changes appear in today's digest)
      this.markLateProjects();

      // 3. Detect and notify status changes (includes Late status changes from step 2)
      this.detectAndNotifyStatusChanges();

      // 4. Sync calendar events (safety net)
      this.syncCalendarEvents();

      console.log('MaintenanceService: Daily maintenance completed');

    } catch (error) {
      console.error(`MaintenanceService: Error during daily maintenance: ${error.message}`);
      this.notificationService.sendErrorNotification(
        'Daily Maintenance Failed',
        `Error: ${error.message}\n\nStack: ${error.stack || 'N/A'}`
      );
      throw error;
    }
  }

  // ===== REMINDERS =====

  /**
   * Sends reminder digest emails for projects with upcoming deadlines.
   * Groups reminders by assignee so each person gets one email with all their reminders.
   */
  sendReminders() {
    DEBUG && console.log('MaintenanceService: Checking reminders');

    const incompleteProjects = this.projectSheet.getIncompleteProjects();

    // Collect reminders grouped by assignee email
    const remindersByAssignee = new Map();

    for (const project of incompleteProjects) {
      // Skip if not in Created status
      if (!project.isCreated) {
        continue;
      }

      const offsets = project.reminderOffsets;
      if (offsets.length === 0) {
        continue;
      }

      const daysUntilDue = project.daysUntilDue(this.today);

      // Check if today matches any reminder offset
      let shouldRemind = false;
      for (const offset of offsets) {
        if (daysUntilDue === offset) {
          shouldRemind = true;
          break;
        }
      }

      if (!shouldRemind) {
        continue;
      }

      DEBUG && console.log(`MaintenanceService: Adding ${daysUntilDue}-day reminder for ${project.projectId}`);

      // Add to each assignee's reminder list
      const assigneeEmails = project.getAssigneeEmails(this.directory);
      for (const email of assigneeEmails) {
        if (!remindersByAssignee.has(email)) {
          remindersByAssignee.set(email, []);
        }
        remindersByAssignee.get(email).push({ project, daysUntilDue });
      }
    }

    // Send digest email to each assignee
    let digestsSent = 0;
    let totalReminders = 0;

    for (const [email, reminders] of remindersByAssignee) {
      this.notificationService.sendReminderDigest(email, reminders);
      digestsSent++;
      totalReminders += reminders.length;
    }

    console.log(`MaintenanceService: Sent ${digestsSent} reminder digest(s) covering ${totalReminders} project(s)`);
  }

  // ===== STATUS CHANGE DETECTION =====

  /**
   * Detects status changes since last snapshot and sends digest emails.
   */
  detectAndNotifyStatusChanges() {
    DEBUG && console.log('MaintenanceService: Detecting status changes');

    // Get current statuses
    const currentStatuses = this.projectSheet.getStatusMap();

    // Initialize snapshot if empty (first run)
    if (this.snapshotSheet.initializeIfEmpty(currentStatuses)) {
      console.log('MaintenanceService: Initialized empty snapshot, skipping change detection');
      return;
    }

    // Detect changes
    const changes = this.snapshotSheet.detectChanges(currentStatuses);

    if (changes.length === 0) {
      DEBUG && console.log('MaintenanceService: No status changes detected');
      this.snapshotSheet.overwriteWithCurrent(currentStatuses);
      return;
    }

    // Build change objects with full project info
    const changeDetails = [];
    for (const change of changes) {
      const project = this.projectSheet.findByProjectId(change.projectId);
      if (project) {
        changeDetails.push({
          project,
          oldStatus: change.oldStatus,
          newStatus: change.newStatus
        });
      }
    }

    // Group changes by recipient (assignees + requested_by)
    const changesByRecipient = this.groupChangesByRecipient(changeDetails);

    // Send digest emails
    for (const [email, recipientChanges] of changesByRecipient) {
      this.notificationService.sendStatusChangeDigest(email, recipientChanges, this.today);
    }

    console.log(`MaintenanceService: Sent status change digest to ${changesByRecipient.size} recipient(s)`);

    // Update snapshot with current statuses
    this.snapshotSheet.overwriteWithCurrent(currentStatuses);
  }

  /**
   * Groups status changes by recipient email.
   * @param {Object[]} changeDetails - Array of {project, oldStatus, newStatus}
   * @returns {Map<string, Object[]>} Map of email -> changes array
   */
  groupChangesByRecipient(changeDetails) {
    const grouped = new Map();

    for (const change of changeDetails) {
      const emails = change.project.getAllRecipientEmails(this.directory);

      for (const email of emails) {
        if (!grouped.has(email)) {
          grouped.set(email, []);
        }
        grouped.get(email).push(change);
      }
    }

    return grouped;
  }

  // ===== LATE PROJECTS =====

  /**
   * Marks projects as late if due date is today and not completed.
   */
  markLateProjects() {
    DEBUG && console.log('MaintenanceService: Checking for late projects');

    const incompleteProjects = this.projectSheet.getIncompleteProjects();
    let marked = 0;

    for (const project of incompleteProjects) {
      // Only process Created projects
      if (!project.isCreated) {
        continue;
      }

      // Check if due today or past due
      if (project.isDueToday(this.today) || project.isPastDue(this.today)) {
        if (project.projectStatus !== PROJECT_STATUS.LATE) {
          project.projectStatus = PROJECT_STATUS.LATE;
          marked++;
          DEBUG && console.log(`MaintenanceService: Marked ${project.projectId} as late`);
        }
      }
    }

    console.log(`MaintenanceService: Marked ${marked} project(s) as late`);
  }

  // ===== CALENDAR SYNC =====

  /**
   * Syncs calendar events with project data (safety net).
   * Catches any discrepancies that may have been missed.
   */
  syncCalendarEvents() {
    DEBUG && console.log('MaintenanceService: Syncing calendar events');

    const createdProjects = this.projectSheet.getCreatedProjects();
    let synced = 0;

    const calendar = CalendarApp.getDefaultCalendar();

    for (const project of createdProjects) {
      if (!project.calendarEventId) {
        continue;
      }

      try {
        const needsSync = this.checkCalendarSync(project, calendar);
        if (needsSync) {
          this.syncCalendarEvent(project, calendar);
          synced++;
        }
      } catch (error) {
        console.warn(`MaintenanceService: Error syncing event for ${project.projectId}: ${error.message}`);
      }
    }

    if (synced > 0) {
      console.log(`MaintenanceService: Synced ${synced} calendar event(s)`);
    }
  }

  /**
   * Checks if a project's calendar event needs syncing.
   * @param {Project} project - The project
   * @param {GoogleAppsScript.Calendar.Calendar} calendar - The calendar
   * @returns {boolean} True if sync is needed
   */
  checkCalendarSync(project, calendar) {
    try {
      const event = withBackoff(() => calendar.getEventById(project.calendarEventId));
      if (!event) {
        return false;
      }

      // Check date
      const eventDate = event.getAllDayStartDate();
      const projectDueDate = project.dueDate;

      if (projectDueDate && !isSameDay(eventDate, projectDueDate)) {
        DEBUG && console.log(`MaintenanceService: Date mismatch for ${project.projectId}`);
        return true;
      }

      // Check attendees
      const eventGuests = event.getGuestList().map(g => g.getEmail().toLowerCase()).sort();
      const projectGuests = project.getAllRecipientEmails(this.directory).map(e => e.toLowerCase()).sort();

      if (JSON.stringify(eventGuests) !== JSON.stringify(projectGuests)) {
        DEBUG && console.log(`MaintenanceService: Attendee mismatch for ${project.projectId}`);
        return true;
      }

      return false;

    } catch (error) {
      DEBUG && console.log(`MaintenanceService: Could not check event for ${project.projectId}: ${error.message}`);
      return false;
    }
  }

  /**
   * Syncs a calendar event with project data.
   * @param {Project} project - The project
   * @param {GoogleAppsScript.Calendar.Calendar} calendar - The calendar
   */
  syncCalendarEvent(project, calendar) {
    try {
      const event = withBackoff(() => calendar.getEventById(project.calendarEventId));
      if (!event) {
        return;
      }

      // Update date
      if (project.dueDate) {
        event.setAllDayDate(project.dueDate);
      }

      // Update title
      event.setTitle(project.displayTitle);

      // Update attendees
      const currentGuests = event.getGuestList().map(g => g.getEmail().toLowerCase());
      const newGuests = project.getAllRecipientEmails(this.directory);

      // Remove guests not in project
      for (const guest of currentGuests) {
        if (!newGuests.map(g => g.toLowerCase()).includes(guest)) {
          event.removeGuest(guest);
        }
      }

      // Add guests not in event
      for (const guest of newGuests) {
        if (!currentGuests.includes(guest.toLowerCase())) {
          try {
            event.addGuest(guest);
          } catch (e) {
            // Ignore individual guest add failures
          }
        }
      }

      // Send notification about the sync
      this.notificationService.sendUpdateNotification(project);

      DEBUG && console.log(`MaintenanceService: Synced calendar event for ${project.projectId}`);

    } catch (error) {
      console.warn(`MaintenanceService: Error syncing event: ${error.message}`);
    }
  }
}

