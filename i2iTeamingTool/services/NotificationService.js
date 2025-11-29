/**
 * NotificationService class - Handles all email notifications.
 * Includes template loading, caching, token substitution, and email sending.
 */
class NotificationService {
  /**
   * Creates a new NotificationService instance.
   * @param {Config} config - The Config instance
   * @param {Directory} directory - The Directory instance
   */
  constructor(config, directory) {
    this.config = config;
    this.directory = directory;
    this.templateCache = new Map();
  }

  // ===== TEMPLATE HANDLING =====

  /**
   * Loads and parses an email template from a Google Doc.
   * First line = subject, remaining lines = body.
   * @param {string} templateDocId - The Google Doc ID
   * @returns {Object} Object with 'subject' and 'body' properties
   */
  loadTemplate(templateDocId) {
    if (!templateDocId) {
      throw new Error('Template Doc ID is required');
    }

    // Check cache first
    if (this.templateCache.has(templateDocId)) {
      DEBUG && console.log(`NotificationService: Using cached template ${templateDocId}`);
      return this.templateCache.get(templateDocId);
    }

    // Load from Google Doc
    const doc = withBackoff(() => DocumentApp.openById(templateDocId));
    const text = doc.getBody().getText();
    const lines = text.split('\n');

    // First non-empty line is the subject
    const subject = lines.find(line => line.trim() !== '') || '';

    // Find where body starts (after first non-empty line)
    const firstLineIndex = lines.findIndex(line => line.trim() !== '');
    const bodyLines = firstLineIndex >= 0 ? lines.slice(firstLineIndex + 1) : [];
    const body = bodyLines.join('\n').trim();

    const template = {
      subject: subject.trim(),
      body: body
    };

    // Cache for reuse
    this.templateCache.set(templateDocId, template);
    DEBUG && console.log(`NotificationService: Loaded template ${templateDocId}`);

    return template;
  }

  /**
   * Prepares an email by loading template and substituting tokens.
   * @param {string} templateDocId - The Google Doc ID
   * @param {Object} tokenValues - Key-value pairs for token substitution
   * @returns {Object} Object with 'subject' and 'body' properties
   */
  prepareEmail(templateDocId, tokenValues) {
    const template = this.loadTemplate(templateDocId);

    return {
      subject: substituteTokens(template.subject, tokenValues),
      body: substituteTokens(template.body, tokenValues)
    };
  }

  /**
   * Sends an email using Gmail.
   * @param {string|string[]} to - Recipient email(s)
   * @param {string} subject - Email subject
   * @param {string} body - Email body (plain text)
   * @param {Object} options - Additional options (cc, bcc, htmlBody, etc.)
   */
  sendEmail(to, subject, body, options = {}) {
    const recipients = Array.isArray(to) ? to.join(',') : to;

    if (!recipients) {
      console.warn('NotificationService.sendEmail: No recipients specified');
      return;
    }

    try {
      withBackoff(() => {
        GmailApp.sendEmail(recipients, subject, body, {
          htmlBody: options.htmlBody || this.textToHtml(body),
          cc: options.cc,
          bcc: options.bcc,
          name: options.name || 'Teaming Tool',
          replyTo: options.replyTo
        });
      });

      console.log(`Sent email to ${recipients}: "${subject}"`);
    } catch (error) {
      console.error(`Failed to send email to ${recipients}: ${error.message}`);
      throw error;
    }
  }

  /**
   * Converts text to basic HTML.
   * TRUSTS the input to contain valid HTML tags.
   * Only converts newlines to <br> and auto-links URLs.
   * @param {string} text - Text with potential HTML tags
   * @returns {string} HTML string
   */
  textToHtml(text) {
    if (!text) return '';

    let html = text;

    // Convert line breaks
    html = html.replace(/\n/g, '<br>');

    // Convert URLs to links (if not already linked)
    // Simple lookahead to avoid double-linking
    html = html.replace(
      /(?<!href=")(https?:\/\/[^\s<]+)/g,
      '<a href="$1">$1</a>'
    );

    return html;
  }

  // ===== NOTIFICATION METHODS =====

  /**
   * Sends new project assignment email to assignees, CC'ing the requester.
   * Sends a single email to all assignees (not individual emails).
   * @param {Project} project - The project
   */
  sendNewProjectEmail(project) {
    const templateId = this.config.emailTemplateNewProject;
    if (!templateId) {
      console.warn('NotificationService: New Project email template not configured');
      return;
    }

    const assigneeEmails = project.getAssigneeEmails(this.directory);
    if (assigneeEmails.length === 0) {
      console.warn(`NotificationService: No assignee emails for project ${project.projectId}`);
      return;
    }

    // Determine greeting name: individual name if one assignee, "All" if multiple
    let assigneeName;
    if (assigneeEmails.length === 1) {
      assigneeName = this.directory.getNameByEmail(assigneeEmails[0]) || assigneeEmails[0];
    } else {
      assigneeName = 'All';
    }

    const tokenValues = project.getTokenValues(this.directory, {
      ASSIGNEE_NAME: assigneeName
    });

    const prepared = this.prepareEmail(templateId, tokenValues);

    // Get requester email for CC
    const requesterEmail = this.directory.resolveToEmail(project.requestedBy);

    // Send single email to all assignees, CC the requester
    this.sendEmail(assigneeEmails, prepared.subject, prepared.body, {
      cc: requesterEmail || undefined
    });
  }

  /**
   * Sends a reminder digest email to a single assignee with all their upcoming project reminders.
   * @param {string} assigneeEmail - The assignee's email
   * @param {Object[]} reminders - Array of {project, daysUntilDue}
   */
  sendReminderDigest(assigneeEmail, reminders) {
    const templateId = this.config.emailTemplateReminder;
    if (!templateId) {
      console.warn('NotificationService: Reminder email template not configured');
      return;
    }

    if (reminders.length === 0) {
      return;
    }

    const assigneeName = this.directory.getNameByEmail(assigneeEmail) || assigneeEmail;

    // If single reminder, use the standard template format
    if (reminders.length === 1) {
      const { project, daysUntilDue } = reminders[0];
      const tokenValues = project.getTokenValues(this.directory, {
        ASSIGNEE_NAME: assigneeName,
        DAYS_UNTIL_DUE: String(daysUntilDue)
      });

      const prepared = this.prepareEmail(templateId, tokenValues);
      this.sendEmail(assigneeEmail, prepared.subject, prepared.body);
      return;
    }

    // Multiple reminders - build a digest
    // Use a custom subject and build the list ourselves
    const subject = `Reminder: ${reminders.length} projects with upcoming deadlines`;

    const remindersList = reminders.map(({ project, daysUntilDue }) => {
      return `• <strong>${project.projectName}</strong> - Due in ${daysUntilDue} days (${formatDate(project.dueDate)})<br>` +
             `  Project ID: ${project.projectId} | <a href="${project.folderUrl}">View Project Folder</a>`;
    }).join('<br><br>');

    const body = `Hello ${assigneeName},<br><br>` +
                 `This is a reminder that the following projects are approaching their deadlines:<br><br>` +
                 `${remindersList}<br><br>` +
                 `Please ensure all work is completed and submitted by the deadlines.<br><br>` +
                 `Thank you.`;

    this.sendEmail(assigneeEmail, subject, body);
  }

  /**
   * Sends status change digest email.
   * @param {string} recipientEmail - Recipient email
   * @param {Object[]} changes - Array of {project, oldStatus, newStatus}
   * @param {Date} date - The date of the digest
   */
  sendStatusChangeDigest(recipientEmail, changes, date) {
    const templateId = this.config.emailTemplateStatusChange;
    if (!templateId) {
      console.warn('NotificationService: Status Change email template not configured');
      return;
    }

    if (changes.length === 0) {
      return;
    }

    const recipientName = this.directory.getNameByEmail(recipientEmail) || recipientEmail;

    // Build the status changes list
    const changesList = changes.map(change => {
      const { project, newStatus } = change;
      return `• <strong>${project.projectName}</strong> - Status changed to: <strong>${newStatus}</strong><br>` +
             `  Project ID: ${project.projectId} | <a href="${project.folderUrl}">View Project Folder</a>`;
    }).join('<br><br>');

    const tokenValues = {
      RECIPIENT_NAME: recipientName,
      DATE: formatDate(date),
      STATUS_CHANGES_LIST: changesList
    };

    const prepared = this.prepareEmail(templateId, tokenValues);
    this.sendEmail(recipientEmail, prepared.subject, prepared.body);
  }

  /**
   * Sends project update notification to assignees, CC'ing the requester.
   * Sends a single email to all assignees (not individual emails).
   * @param {Project} project - The updated project
   */
  sendUpdateNotification(project) {
    const templateId = this.config.emailTemplateUpdate;
    if (!templateId) {
      console.warn('NotificationService: Project Update email template not configured');
      return;
    }

    const assigneeEmails = project.getAssigneeEmails(this.directory);
    if (assigneeEmails.length === 0) {
      console.warn(`NotificationService: No assignee emails for project ${project.projectId}`);
      return;
    }

    // Determine greeting name: individual name if one assignee, "All" if multiple
    let recipientName;
    if (assigneeEmails.length === 1) {
      recipientName = this.directory.getNameByEmail(assigneeEmails[0]) || assigneeEmails[0];
    } else {
      recipientName = 'All';
    }

    const tokenValues = project.getTokenValues(this.directory, {
      RECIPIENT_NAME: recipientName
    });

    const prepared = this.prepareEmail(templateId, tokenValues);

    // Get requester email for CC
    const requesterEmail = this.directory.resolveToEmail(project.requestedBy);

    // Send single email to all assignees, CC the requester
    this.sendEmail(assigneeEmails, prepared.subject, prepared.body, {
      cc: requesterEmail || undefined
    });
  }

  /**
   * Sends project cancellation notification to assignees, CC'ing the requester.
   * Sends a single email to all assignees (not individual emails).
   * @param {Project} project - The cancelled project
   */
  sendCancellationNotification(project) {
    const templateId = this.config.emailTemplateCancellation;
    if (!templateId) {
      console.warn('NotificationService: Project Cancellation email template not configured');
      return;
    }

    const assigneeEmails = project.getAssigneeEmails(this.directory);
    if (assigneeEmails.length === 0) {
      console.warn(`NotificationService: No assignee emails for project ${project.projectId}`);
      return;
    }

    // Determine greeting name: individual name if one assignee, "All" if multiple
    let recipientName;
    if (assigneeEmails.length === 1) {
      recipientName = this.directory.getNameByEmail(assigneeEmails[0]) || assigneeEmails[0];
    } else {
      recipientName = 'All';
    }

    const tokenValues = project.getTokenValues(this.directory, {
      RECIPIENT_NAME: recipientName
    });

    const prepared = this.prepareEmail(templateId, tokenValues);

    // Get requester email for CC
    const requesterEmail = this.directory.resolveToEmail(project.requestedBy);

    // Send single email to all assignees, CC the requester
    this.sendEmail(assigneeEmails, prepared.subject, prepared.body, {
      cc: requesterEmail || undefined
    });
  }

  /**
   * Sends error notification to configured admin emails.
   * @param {string} subject - Error subject
   * @param {string} message - Error message/details
   * @param {Object} options - Optional settings
   * @param {string} options.cc - Email address to CC (e.g., the requester)
   */
  sendErrorNotification(subject, message, options = {}) {
    const adminEmails = this.config.errorEmailAddresses;
    if (adminEmails.length === 0) {
      console.warn('NotificationService: No admin emails configured for error notifications');
      return;
    }

    const fullSubject = `[Teaming Tool Error] ${subject}`;
    const body = `An error occurred in the Teaming Tool automation:\n\n${message}\n\n` +
                 `Time: ${new Date().toLocaleString()}`;

    // CC is best-effort - if provided and valid, include it
    const emailOptions = {};
    if (options.cc && isValidEmail(options.cc)) {
      emailOptions.cc = options.cc;
    }

    this.sendEmail(adminEmails, fullSubject, body, emailOptions);
  }

  /**
   * Clears the template cache.
   */
  clearCache() {
    this.templateCache.clear();
    DEBUG && console.log('NotificationService: Cleared template cache');
  }
}

