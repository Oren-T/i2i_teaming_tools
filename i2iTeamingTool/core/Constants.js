/**
 * Constants for the Project Management Tool library.
 * Contains required keys, status enums, and sheet names.
 */

// ===== DEBUG FLAG =====
// Default is false. Can be enabled per-district by adding "Debug Mode" = "true" to Config sheet.
// ExecutionContext sets this at runtime based on Config sheet value.
let DEBUG = false;

// ===== PRODUCT METADATA =====
const PRODUCT_NAME = 'Project Management Tool';
const PRODUCT_ERROR_PREFIX = '[Project Management Tool Error]';

// ===== SHEET NAMES =====
const SHEET_NAMES = {
  PROJECTS: 'Project Management Sheet',
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
// Core statuses required for system logic (e.g. completion tracking, late detection).
// Other statuses are flexible and read from the Codes sheet.
const PROJECT_STATUS = {
  PROJECT_ASSIGNED: 'Project Assigned',
  ON_TRACK: 'On Track',
  BEHIND_SCHEDULE: 'Behind Schedule',
  STUCK: 'Stuck',
  LATE: 'Late',
  COMPLETE: 'Complete'
};

// ===== REQUIRED CONFIG KEYS =====
// These keys must exist in the Config sheet (Column A)
const REQUIRED_CONFIG_KEYS = [
  'District ID',
  'Next Serial',
  'Parent Folder ID',
  'Root Folder ID',
  'Main Spreadsheet ID',
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
  'Debug Mode',              // Set to "true" to enable verbose logging
  'School Year Start Month', // Month (1-12) when school year begins (default: 7 for July)
  'Backups Folder ID'        // Google Drive folder ID where weekly backups are stored
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
  'file_id',
  'notes'
];

// ===== DIRECTORY COLUMNS =====
const DIRECTORY_COLUMNS = {
  NAME: 'Name',
  EMAIL: 'Email Address',
  PERMISSIONS: 'Permissions',      // Legacy (pre-permission-model overhaul)
  ACTIVE: 'Active?',
  GLOBAL_ACCESS: 'Global Access',
  MAIN_FILE_ROLE: 'Project Directory Role',
  PROJECT_SCOPE: 'Project Folders Role'
};

// ===== DIRECTORY ACCESS ROLES & SCOPES =====
// Normalized internal values used by permission evaluation logic.
const DIRECTORY_ACCESS_ROLES = {
  NONE: 'none',
  VIEWER: 'viewer',
  EDITOR: 'editor'
};

const DIRECTORY_FOLDER_SCOPES = {
  NONE: 'none',
  ASSIGNED_ONLY: 'assigned_only',
  ALL_VIEWER: 'all_viewer',
  ALL_EDITOR: 'all_editor'
};

// ===== CODES COLUMNS =====
const CODES_COLUMNS = {
  CATEGORY: 'Category',
  STATUS: 'Status',
  REMINDER_DAYS_OFFSET: 'Reminder Days',
  REMINDER_DAYS_READABLE: 'Reminder Days: Readable'
};

// ===== CODES SHEET LAYOUT =====
const CODES_LAYOUT = {
  HEADER_ROW: 3,
  CATEGORY_COL: 1,          // Column A
  STATUS_COL: 3,            // Column C
  REMINDER_OFFSET_COL: 5,   // Column E
  REMINDER_LABEL_COL: 6     // Column F
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
// Multiple titles can map to the same internal key for backward compatibility
const FORM_FIELD_MAP = {
  'Goal #': 'goal_number',
  'Goal # (if available)': 'goal_number',
  'LCAP Goal # (if available)': 'goal_number',
  'Action #': 'action_number',
  'Action # (if available)': 'action_number',
  'LCAP Action # (if available)': 'action_number',
  'Category': 'category',
  'Title': 'project_name',              // Legacy title
  'Project Title': 'project_name',       // Preferred title
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

// ===== CALENDAR EVENT COLORS =====
// Maps project status to Google Calendar EventColor IDs
// See: https://developers.google.com/apps-script/reference/calendar/event-color
const CALENDAR_COLORS = {
  // Status -> color mapping
  [PROJECT_STATUS.PROJECT_ASSIGNED]: CalendarApp.EventColor.GRAY,
  [PROJECT_STATUS.COMPLETE]: CalendarApp.EventColor.GREEN,
  [PROJECT_STATUS.LATE]: CalendarApp.EventColor.RED,
  // Default for all other statuses (On Track, Behind Schedule, Stuck, etc.)
  DEFAULT: CalendarApp.EventColor.YELLOW
};

/**
 * Gets the calendar color for a project status.
 * @param {string} status - The project status
 * @returns {CalendarApp.EventColor} The calendar color
 */
function getCalendarColorForStatus(status) {
  if (!status) {
    return CALENDAR_COLORS.DEFAULT;
  }
  return CALENDAR_COLORS[status] || CALENDAR_COLORS.DEFAULT;
}

