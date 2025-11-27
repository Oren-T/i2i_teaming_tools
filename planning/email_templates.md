# Email Templates

Draft text for automated email notifications. Use token substitution for placeholders.

**Structure:** Each template is stored as a separate Google Doc. First line = subject line, remaining lines = email body. Apps Script parses by splitting on newlines.

---

## New Project Assignment

New Project Assigned: {{PROJECT_TITLE}}

Hello {{ASSIGNEE_NAME}},

A new project has been assigned to you:

**Project:** {{PROJECT_TITLE}}  
**Category:** {{CATEGORY}}  
**Requested by:** {{REQUESTED_BY_NAME}}  
**Deadline:** {{DEADLINE}}  
**Project ID:** {{PROJECT_ID}}

**Description:**  
{{DESCRIPTION}}

You can access the project folder and all related materials here: {{FOLDER_LINK}}

A calendar event has been created for this deadline. Please review the project details and let us know if you have any questions.

Thank you.

---

## Reminder Notification

Reminder: {{PROJECT_TITLE}} - Due in {{DAYS_UNTIL_DUE}} days

Hello {{ASSIGNEE_NAME}},

This is a reminder that your project is approaching its deadline:

**Project:** {{PROJECT_TITLE}}  
**Due Date:** {{DEADLINE}}  
**Days Remaining:** {{DAYS_UNTIL_DUE}}  
**Project ID:** {{PROJECT_ID}}

Access the project folder: {{FOLDER_LINK}}

Please ensure all work is completed and submitted by the deadline.

Thank you.

---

## Status Change Digest

Project Status Updates - {{DATE}}

Hello {{RECIPIENT_NAME}},

The following projects have had status changes:

{{STATUS_CHANGES_LIST}}

Where:
* **{{PROJECT_TITLE}}** - Status changed to: **{{NEW_STATUS}}**  
  Project ID: {{PROJECT_ID}} | [View Project Folder]({{FOLDER_LINK}})

---

If you have any questions about these updates, please contact the project leads directly.

Thank you.

---

## Token Reference

| Token | Description |
|-------|-------------|
| `{{PROJECT_TITLE}}` | Project name/title |
| `{{PROJECT_ID}}` | Unique project identifier (e.g., NUSD-25_26-0024) |
| `{{ASSIGNEE_NAME}}` | Name of person assigned to project |
| `{{REQUESTED_BY_NAME}}` | Name of person who requested the project |
| `{{CATEGORY}}` | Project category (e.g., LCAP, SPSA) |
| `{{DEADLINE}}` | Project deadline date |
| `{{DAYS_UNTIL_DUE}}` | Number of days until deadline |
| `{{DESCRIPTION}}` | Project description |
| `{{FOLDER_LINK}}` | URL to project folder |
| `{{NEW_STATUS}}` | Updated project status |
| `{{DATE}}` | Date of status change digest |
| `{{RECIPIENT_NAME}}` | Name of digest recipient |
| `{{STATUS_CHANGES_LIST}}` | Formatted list of all status changes for digest |

---

## Parsing Function

Apps Script function to parse email templates from Google Docs; note, `sendEmail` should use HTML content rendering for full support of formatting:

```javascript
/**
 * Parses an email template from a Google Doc
 * @param {string} templateDocId - Google Doc ID containing the template
 * @returns {Object} Object with 'subject' and 'body' properties
 */
function getEmailTemplate(templateDocId) {
  const doc = DocumentApp.openById(templateDocId);
  const text = doc.getBody().getText();
  const lines = text.split('\n');
  
  // First non-empty line is the subject
  const subject = lines.find(line => line.trim() !== '') || '';
  
  // Find where body starts (after first non-empty line)
  const firstLineIndex = lines.findIndex(line => line.trim() !== '');
  const bodyLines = lines.slice(firstLineIndex + 1);
  const body = bodyLines.join('\n').trim();
  
  return { subject: subject.trim(), body: body };
}
```
