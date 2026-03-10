// ============================================================
//  CONFIGURATION — Department Task Tracker
// ============================================================
//  Review and update these constants before first deployment.
//  See: department_task_tracker_detailed_spec.md
// ============================================================

// ===== Error Notification Recipients =====
const ERROR_EMAILS = [
  'orendev4@gmail.com',
  'joel@inform2inspire.com'
];

// ===== Reminder Schedule =====
// Days before a task deadline to send reminder emails.
const REMINDER_DAYS = [3, 7];

// ===== Calendar =====
const CALENDAR_ID = 'primary';

// ===== Tab Names =====
// Must match the spreadsheet template exactly.
const TAB_NAMES = {
  PROJECTS: 'Projects / Initiatives',
  TASKS: 'Tasks',
  DIRECTORY: 'Directory'
};

// ===== Column Headers =====
// Must match Row 1 of each tab exactly.

const PROJECT_COLUMNS = {
  TITLE: 'Title',
  LEADER: 'Leader',
  DEADLINE: 'Deadline',
  STATUS: 'Status',
  LINKS: 'Links to Folder'
};

const TASK_COLUMNS = {
  TASK_NAME: 'Task Name',
  PROJECT: 'Project / Initiative',
  ASSIGNEE: 'Assignee',
  DEADLINE: 'Deadline',
  DOCUMENTS: 'Document & Links',
  NOTES: 'Notes',
  STATUS: 'Status',
  UPDATE_REMINDERS: 'Update Reminders',
  CALENDAR_EVENT_ID: '_calendar_event_id'
};

const DIRECTORY_COLUMNS = {
  NAME: 'Name',
  EMAIL: 'Email',
  NOTES: 'Notes'
};

// ===== Status Values =====

const TASK_STATUS = {
  NOT_STARTED: 'Not Started',
  IN_PROGRESS: 'In Progress',
  BEHIND_SCHEDULE: 'Behind Schedule',
  COMPLETE: 'Complete'
};

const PROJECT_STATUS = {
  ACTIVE: 'Active',
  COMPLETE: 'Complete',
  CANCELLED: 'Cancelled'
};

// ===== Email Templates =====
// {{TOKEN}} placeholders are replaced at send time via substituteTokens().

const EMAIL_TEMPLATES = {
  REMINDER: {
    SUBJECT: 'Reminder: {{TASK_NAME}} \u2014 Due in {{DAYS_UNTIL_DUE}} days',
    BODY: [
      '<div style="font-family: Arial, sans-serif; max-width: 600px;">',
      '  <h2 style="color: #333;">Task Reminder</h2>',
      '  <p>This is a reminder that the following task is due soon:</p>',
      '  <table style="border-collapse: collapse; width: 100%; margin: 16px 0;">',
      '    <tr>',
      '      <td style="padding: 8px 12px; font-weight: bold; background: #f5f5f5;">Task</td>',
      '      <td style="padding: 8px 12px; background: #f5f5f5;">{{TASK_NAME}}</td>',
      '    </tr>',
      '    <tr>',
      '      <td style="padding: 8px 12px; font-weight: bold;">Project</td>',
      '      <td style="padding: 8px 12px;">{{PROJECT_NAME}}</td>',
      '    </tr>',
      '    <tr>',
      '      <td style="padding: 8px 12px; font-weight: bold; background: #f5f5f5;">Deadline</td>',
      '      <td style="padding: 8px 12px; background: #f5f5f5;">{{DEADLINE}}</td>',
      '    </tr>',
      '    <tr>',
      '      <td style="padding: 8px 12px; font-weight: bold;">Days Remaining</td>',
      '      <td style="padding: 8px 12px;">{{DAYS_UNTIL_DUE}}</td>',
      '    </tr>',
      '  </table>',
      '</div>'
    ].join('\n')
  },

  ERROR: {
    SUBJECT: '[Task Tracker Error] Setup issue detected',
    BODY: [
      '<div style="font-family: Arial, sans-serif; max-width: 600px;">',
      '  <h2 style="color: #cc0000;">Task Tracker Error</h2>',
      '  <p>A critical error was detected in the Department Task Tracker:</p>',
      '  <pre style="background: #f5f5f5; padding: 12px; border-left: 4px solid #cc0000;">{{ERROR_MESSAGE}}</pre>',
      '  <p>Spreadsheet: <a href="{{SPREADSHEET_URL}}">Open Spreadsheet</a></p>',
      '</div>'
    ].join('\n')
  }
};
