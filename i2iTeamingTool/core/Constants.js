/**
 * Constants for the i2i Teaming Tool library.
 * Contains required keys, status enums, and sheet names.
 */

// ===== DEBUG FLAG =====
// Default is false. Can be enabled per-district by adding "Debug Mode" = "true" to Config sheet.
// ExecutionContext sets this at runtime based on Config sheet value.
let DEBUG = false;

// ===== SHEET NAMES =====
const SHEET_NAMES = {
  PROJECTS: 'Projects',
  CONFIG: 'Config',
  DIRECTORY: 'Directory',
  CODES: 'Codes',
  STATUS_SNAPSHOT: 'Status Snapshot',
  FORM_RESPONSES: 'Form Responses (Raw)'
};

// ===== AUTOMATION STATUS VALUES =====
const AUTOMATION_STATUS = {
  BLANK: '',
  READY: 'Ready',
  CREATED: 'Created',
  UPDATED: 'Updated',
  DELETE_NOTIFY: 'Delete (Notify)',
  DELETE_NO_NOTIFY: 'Delete (Don\'t Notify)',
  DELETED: 'Deleted',
  ERROR: 'Error'
};

// ===== PROJECT STATUS VALUES =====
const PROJECT_STATUS = {
  NOT_STARTED: 'Not started',
  BEHIND_SCHEDULE: 'Behind schedule',
  STUCK: 'Stuck',
  ON_TRACK: 'On track',
  COMPLETED: 'Completed',
  LATE: 'Late'
};

// ===== REQUIRED CONFIG KEYS =====
// These keys must exist in the Config sheet (Column A)
const REQUIRED_CONFIG_KEYS = [
  'District ID',
  'School Year',
  'Next Serial',
  'Parent Folder ID',
  'Project Template ID',
  'Form ID',
  'Error Email Addresses',
  'Email Template - New Project',
  'Email Template - Reminder',
  'Email Template - Status Change',
  'Email Template - Project Update',
  'Email Template - Project Cancellation'
];

// ===== OPTIONAL CONFIG KEYS =====
// These keys are optional in the Config sheet
const OPTIONAL_CONFIG_KEYS = [
  'Debug Mode'  // Set to "true" to enable verbose logging
];

// ===== REQUIRED PROJECT COLUMNS =====
// These internal keys must exist in Row 2 of the Projects sheet
const REQUIRED_PROJECT_COLUMNS = [
  'project_id',
  'created_at',
  'school_year',
  'goal_number',
  'action_number',
  'category',
  'project_name',
  'description',
  'assignee',
  'requested_by',
  'due_date',
  'project_status',
  'completed_at',
  'reminder_offsets',
  'automation_status',
  'calendar_event_id',
  'folder_id',
  'notes'
];

// ===== DIRECTORY COLUMNS =====
const DIRECTORY_COLUMNS = {
  NAME: 'Name',
  EMAIL: 'Email Address',
  PERMISSIONS: 'Permissions'
};

// ===== CODES COLUMNS =====
const CODES_COLUMNS = {
  CATEGORY: 'Category',
  STATUS: 'Status',
  REMINDER_DAYS_OFFSET: 'Reminder Days Offset',
  REMINDER_DAYS_READABLE: 'Reminder Days: Readable'
};

// ===== EMAIL TEMPLATE TOKENS =====
const EMAIL_TOKENS = {
  PROJECT_TITLE: '{{PROJECT_TITLE}}',
  PROJECT_ID: '{{PROJECT_ID}}',
  ASSIGNEE_NAME: '{{ASSIGNEE_NAME}}',
  REQUESTED_BY_NAME: '{{REQUESTED_BY_NAME}}',
  CATEGORY: '{{CATEGORY}}',
  DEADLINE: '{{DEADLINE}}',
  DAYS_UNTIL_DUE: '{{DAYS_UNTIL_DUE}}',
  DESCRIPTION: '{{DESCRIPTION}}',
  FOLDER_LINK: '{{FOLDER_LINK}}',
  NEW_STATUS: '{{NEW_STATUS}}',
  DATE: '{{DATE}}',
  RECIPIENT_NAME: '{{RECIPIENT_NAME}}',
  STATUS_CHANGES_LIST: '{{STATUS_CHANGES_LIST}}'
};

// ===== FORM FIELD NAMES =====
// Maps form question titles to internal keys
const FORM_FIELD_MAP = {
  'Goal #': 'goal_number',
  'Action #': 'action_number',
  'Category': 'category',
  'Title': 'project_name',
  'Description': 'description',
  'Assigned to': 'assignee',
  'Deadline': 'due_date',
  'Notes': 'notes'
};

// ===== DEFAULT VALUES =====
const DEFAULTS = {
  CATEGORY: 'LCAP',
  REMINDER_OFFSETS: [3, 7, 14]
};

