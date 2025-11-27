# Code Ideas

> **Note:** Everything in this document represents *suggestions and ideas* for implementation, not hard requirements. These patterns emerged from reviewing the old code and thinking through architecture, but the actual implementation may differ based on practical constraints, testing, or better ideas that emerge during development. Treat this as a starting point for discussion, not gospel.

---

## Interesting Old Logic

### Document Properties for Config
The old code used `PropertiesService.getDocumentProperties()` to store persistent configuration:
- File IDs (Directory, Form, Template, Project Folders, etc.)
- Project ID counter (auto-incremented serial number)
- Error email list

This allowed configuration to survive across script executions without hardcoding values. Files were discovered by name in the parent folder during first-time setup, then their IDs were persisted.

**Note for new implementation:** We're using a Config sheet tab instead of Document Properties, which is more user-visible and editable. The Config sheet stores District ID, School Year, Next Serial, Parent Folder ID, Template IDs, Form ID, Email Template Doc IDs, etc.

### Assertion / Guard Functions
Pattern for validating that required globals are initialized before use:

```javascript
function Directory_assertions() {
  if (Directory == null) {
    console.error("Directory variables not yet populated");
    throw "Unable to access the directory file properly; please contact Joel Rabin.";
  }
  console.log("Directory variables populated; attempting directory query");
}

function StaffSettings_assertions() {
  if (StaffSettingsSheet == null) {
    console.error("Staff Settings Sheet variables not yet populated");
    GmailApp.sendEmail(notifyForErrors, "Teaming Tool Error", "...");
    throw "Unable to access the staff settings file properly; please contact Joel Rabin.";
  }
  console.log("Staff settings variables populated; attempting query");
}
```

These provide early failure with clear error messages when initialization is missed.

### First-Time Setup with UI Prompts
Interactive setup flow that:
1. Creates a custom menu ("Teaming Tool" > "First-Time Setup")
2. Discovers required files by name in the parent folder
3. Validates single file exists for each expected name
4. Alerts user if files are missing or duplicated
5. Stores discovered file IDs in Document Properties
6. Creates required triggers programmatically
7. Guides user through error email list configuration with prompts

This made deployment to new districts self-service rather than requiring manual ID configuration.

---

## Patterns to Carry Forward

### Dynamic Column Lookup by Header Name
Rather than hardcoding column indices, read headers and build a `key → column index` map at runtime. This makes the system resilient to column reordering:

```javascript
function columnNumberFromName(sheetVals, colName) {
  var topRow = sheetVals[0];
  if (topRow.indexOf(colName) != -1) {
    return topRow.indexOf(colName);
  } else {
    // Error handling with admin notification
    throw ("Error: Invalid column name");
  }
}
```

The new implementation will use Row 2 as internal keys and validate against `EXPECTED_KEYS`.

### Domain-Based File Organization
Logical separation by responsibility:
- Directory management
- Project lifecycle (creation, updates, deletion)
- Notifications (emails, calendar)
- Configuration and utilities

### Central Library + Thin Client Deployment
All business logic lives in a shared library. Each district has a minimal client script with local config that delegates to the library. Updates can be pushed to all districts by updating the library version.

**Architecture decision: Single bound script per district.**
- The Google Form is configured to send responses to a **hidden tab** ("Form Responses (Raw)") within the Main Projects File spreadsheet.
- This allows a single spreadsheet-bound script to handle *all* triggers: `onFormSubmit`, `onEdit`, time-driven batch/daily, and manual menu actions.
- **Why:** Avoids needing two separate script projects (form-bound + spreadsheet-bound) per district, simplifying deployment, library references, and trigger management.

### Admin Error Notifications
Configurable list of email addresses that receive error notifications when automation fails. Provides observability without requiring log access.

### Row-by-Row Processing with Active Context
Pattern of iterating through rows and setting "active" context variables for the current row being processed, then writing back after modifications.

---

## Proposed Class Architecture

### Design Principles

1. **Replace mutable global state with dependency injection** — Services receive what they need via constructor parameters rather than reaching into global variables. This makes dependencies explicit and testable.

2. **Encapsulate domain logic into classes** — Group related operations by *what they do*, not *when they run*. A `ProjectService` owns all project lifecycle logic; trigger handlers become thin dispatchers.

3. **Config schema validation upfront** — Validate all required config keys and column headers at startup, fail fast with a complete error list before any mutations happen.

4. **Cache sheet data in class instances** — Load data once, work in memory, write back at the end. Minimizes expensive `getValue()`/`setValue()` API calls.

### Suggested Classes

| Layer | Class | Responsibility |
|-------|-------|----------------|
| Context | `ExecutionContext` | Bundles all cached state for a single run; constructed once at entry point |
| Data | `Config` | Config sheet access, typed getters |
| Data | `IdAllocator` | Lock-protected serial generation for project IDs |
| Data | `ProjectSheet` | Main projects data, column indexing, creates/manages `Project` instances |
| Data | `Project` | Row-level wrapper with typed accessors, computed properties, dirty tracking |
| Data | `SnapshotSheet` | Status snapshot for daily change detection |
| Data | `Directory` | Staff name ↔ email lookups, active staff list |
| Data | `Codes` | Dropdown values (categories, statuses, reminder offsets) |
| Service | `ProjectService` | Create/update/delete project assets (folders, templates, calendar events) |
| Service | `NotificationService` | All email sending, template loading/caching from Google Docs, token substitution |
| Service | `MaintenanceService` | Daily tasks: reminders, status-change digest, late marking, calendar sync |
| Service | `FormService` | Sync Google Form dropdowns with Directory/Codes sheets |
| Utility | `Validator` | Startup validation of config keys, project columns, file access |

### ExecutionContext

**`ExecutionContext`** — Bundles all cached state for a single script execution. Constructed once at the entry point, passed to all services.

```javascript
class ExecutionContext {
  constructor(spreadsheetId) {
    this.ss = SpreadsheetApp.openById(spreadsheetId);
    this.now = new Date();  // Consistent timestamp for entire run
    
    // Data layer
    this.config = new Config(this.ss.getSheetByName('Config'));
    this.idAllocator = new IdAllocator(this.config);
    this.projectSheet = new ProjectSheet(this.ss.getSheetByName('Projects'));
    this.snapshotSheet = new SnapshotSheet(this.ss.getSheetByName('Status Snapshot'));
    this.directory = new Directory(this.ss.getSheetByName('Directory'));
    this.codes = new Codes(this.ss.getSheetByName('Codes'));
    
    // Utility
    this.validator = new Validator(this.config, this.projectSheet);
    
    // Services
    this.notificationService = new NotificationService(this.config, this.directory);
    this.projectService = new ProjectService(this);
    this.maintenanceService = new MaintenanceService(this);
    this.formService = new FormService(this);
  }
}
```

This makes "bundle of cached state for one execution" a first-class concept. Services receive the context (or specific dependencies from it) rather than reaching into globals.

### Data Layer Classes

**`Config`** — Wraps the Config sheet tab.
- Typed getters: `districtId`, `schoolYear`, `parentFolderId`, `projectTemplateId`, `formId`, `errorEmailAddresses`, `emailTemplateIds`
- `getAndIncrementSerial()` — Reads and increments the serial value (called by `IdAllocator`)

**`IdAllocator`** — Lock-protected project ID generation.
- Separates the locking concern from config reading
- Formats IDs as `DIST-yy_yy-####`

```javascript
class IdAllocator {
  constructor(config) {
    this.config = config;
  }
  
  next() {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const serial = this.config.getAndIncrementSerial();
      const id = `${this.config.districtId}-${this.config.schoolYear}-${String(serial).padStart(4, '0')}`;
      return id;
    } finally {
      lock.releaseLock();
    }
  }
}
```

**`ProjectSheet`** — Wraps the main Projects sheet.
- `loadData()` — Single API call to load all data, creates `Project` instances
- `getProjects()` / `getProjectsWhere(predicate)` — Access cached Project objects
- `getReadyProjects()`, `getPendingDeleteProjects()`, `getUpdatedProjects()` — Convenience filters
- `getColumnIndex(key)` — Lookup by Row 2 internal keys
- `flush()` — Writes only dirty Project rows back to sheet

**`Project`** — Row-level abstraction with typed accessors and dirty tracking.
- **Core:** `get(key)`, `set(key, value)`, `isDirty()`, `getDirtyEntries()`
- **Typed getters:** `projectId`, `projectName`, `dueDate` (returns Date), `assignee` (returns string, may be comma-separated for multiple), `reminderOffsets` (returns number array), etc.
- **Typed setters:** `projectId`, `automationStatus`, `folderId`, `calendarEventId`, `completedAt`, etc.
- **Computed:** `folderUrl`, `displayTitle` (`[ID] Name`), `daysUntilDue(refDate)`
- **Predicates:** `isReady`, `isComplete`, `isLate`, `isPendingDelete`, `shouldNotifyOnDelete`, `isDueForReminder(offset, refDate)`

This transforms service code from:
```javascript
const colIdx = sheet.getColumnIndex('automation_status');
if (row[colIdx] === 'Ready') { ... }
row[sheet.getColumnIndex('project_id')] = id;
```

To:
```javascript
if (project.isReady) { ... }
project.projectId = id;
```

Dirty tracking enables efficient partial writes — only modified cells are written back on `flush()`.

**`SnapshotSheet`** — Wraps the Status Snapshot sheet.
- `loadSnapshot()` — Returns `Map<projectId, status>`
- `overwriteWithCurrent(projectStatusMap)` — Replace snapshot after digest sent

**`Directory`** — Wraps the Directory sheet.
- `getNameByEmail(email)` / `getEmailByName(name)`
- `getActiveStaffNames()` / `getActiveStaffEmails()` — For form dropdown sync

**`Codes`** — Wraps the Codes sheet.
- `getCategories()`, `getStatuses()`, `getReminderOffsets()`

### Service Layer Classes

**`ProjectService`** — Handles the 10-minute batch trigger.
- `processReadyProjects()` — Generate ID, create folder, copy templates, create calendar event, send email, set status to Created
- `processUpdatedProjects()` — Re-sync calendar event, send update notifications
- `processDeleteRequests()` — Cancel calendar event, optionally notify, hide row

**`NotificationService`** — All email sending.
- `sendNewProjectEmail(row)`
- `sendReminderEmail(row, daysUntilDue)`
- `sendStatusChangeDigest(recipient, changes)`
- `sendUpdateNotification(row)` / `sendCancellationNotification(row)`
- Internal: `loadTemplate(docId)` (caches templates in a Map for reuse within the run), `substituteTokens(template, values)`

**`MaintenanceService`** — Handles the daily 8am trigger.
- `runDailyMaintenance()` — Orchestrates all daily tasks
- Internal: `sendReminders()`, `detectAndNotifyStatusChanges()`, `markLateProjects()`, `syncCalendarEvents()`

**`FormService`** — Handles form dropdown sync.
- `syncAllDropdowns()`
- `syncAssigneeDropdown()` / `syncCategoryDropdown()`

### Utility

**`Validator`** — Validates environment at startup.
- `validate()` — Throws with collected errors if anything is invalid
- `validateConfigKeys()` — Check all required keys exist in Config sheet
- `validateProjectColumns()` — Check all required keys exist in Row 2, no duplicates
- `validateFileAccess()` — Verify we can access Parent Folder, Template, Form, etc.

**Constants and helper functions:**
- `REQUIRED_CONFIG_KEYS` — Array of expected config keys
- `REQUIRED_PROJECT_COLUMNS` — Array of expected Row 2 column keys
- `AUTOMATION_STATUS` — Enum-like object `{ READY: 'Ready', CREATED: 'Created', UPDATED: 'Updated', ... }`
- Pure functions: `formatProjectId()`, `parseReminderOffsets()`, `formatDate()`, etc.

---

## Apps Script Library Integration

### The Challenge with Classes in Libraries

When you add a library with identifier `TeamingToolLib`, you call into it as `TeamingToolLib.someFunction()`. Only **enumerable global properties** are visible to consumers — function declarations, `var` globals, and things explicitly attached to the global object.

A bare `class ProjectService { ... }` declaration creates a **non-enumerable** binding, so it does *not* show up as `TeamingToolLib.ProjectService`. This means `new TeamingToolLib.ProjectService(...)` will throw.

**However**, you *can* expose a class constructor by binding it to an enumerable global:

```javascript
// In the library - this WILL be accessible
var ProjectService = class ProjectService {
  constructor(config) { /* ... */ }
};
```

Then in the client: `const svc = new TeamingToolLib.ProjectService(config);` works.

### Recommended Approach: High-Level Entry Points

Rather than exposing classes or factory functions, expose a handful of top-level orchestrator functions that handle everything internally:

**In the Library:**

```javascript
// ===== PUBLIC API (exposed to client scripts) =====

function processNewProjects(spreadsheetId) {
  const ctx = new ExecutionContext(spreadsheetId);
  ctx.validator.validate();
  ctx.projectService.processReadyProjects();
  ctx.projectService.processUpdatedProjects();
  ctx.projectService.processDeleteRequests();
  ctx.projectSheet.flush();
}

function runDailyMaintenance(spreadsheetId) {
  const ctx = new ExecutionContext(spreadsheetId);
  ctx.validator.validate();
  ctx.maintenanceService.runDailyMaintenance();
}

function syncFormDropdowns(spreadsheetId) {
  const ctx = new ExecutionContext(spreadsheetId);
  ctx.formService.syncAllDropdowns();
}

// Called by spreadsheet-bound onFormSubmit trigger
// Event object contains: e.values, e.namedValues, e.range (row in hidden Form Responses tab)
function handleFormSubmission(spreadsheetId, event) {
  const ctx = new ExecutionContext(spreadsheetId);
  ctx.validator.validate();
  ctx.projectService.normalizeAndAppendFormResponse(event);
  ctx.projectService.processReadyProjects();
  ctx.projectSheet.flush();
}
```

The `ExecutionContext` class (defined in the Data Layer section above) handles all the wiring internally.

**In the Client Script (per district):**

```javascript
// Extremely thin - just config and trigger wiring
const SPREADSHEET_ID = 'abc123...';  // This district's Main Projects File

function onBatchTrigger() {
  TeamingToolLib.processNewProjects(SPREADSHEET_ID);
}

function onDailyTrigger() {
  TeamingToolLib.runDailyMaintenance(SPREADSHEET_ID);
}

function onFormSubmit(e) {
  // e.values, e.namedValues, e.range available from spreadsheet-bound trigger
  TeamingToolLib.handleFormSubmission(SPREADSHEET_ID, e);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Teaming Tool')
    .addItem('Run Now', 'manualRun')
    .addItem('Sync Form Dropdowns', 'syncDropdowns')
    .addItem('Refresh Permissions', 'refreshPermissions')
    .addToUi();
}

function manualRun() {
  TeamingToolLib.processNewProjects(SPREADSHEET_ID);
}

function syncDropdowns() {
  TeamingToolLib.syncFormDropdowns(SPREADSHEET_ID);
}
```

### Benefits of This Approach

| Aspect | Benefit |
|--------|---------|
| Client simplicity | Client scripts are trivial — just config values and trigger wiring |
| Update propagation | Update the library version, all districts get changes automatically |
| Internal flexibility | Classes and logic stay internal; refactor freely without breaking clients |
| Testing | Can test library functions in isolation; mock `buildContext()` for unit tests |
| Clear contract | Public API is just 4-5 functions with obvious names |

### BackoffClient (brief)

- Purpose: centralize retries with exponential backoff + jitter for transient Apps Script service errors (Drive, Calendar, Gmail).
- Defaults: attempts=5, baseDelayMs=250, maxDelayMs=5000, full jitter; retries on common transient signals (HTTP 429/5xx, "Service invoked too many times", "Rate Limit Exceeded", "Service unavailable").
- Usage: wrap idempotent operations only; log attempt count and last error; bubble up final error for observability.

```javascript
// Pseudocode signature
function withBackoff(fn, {
  attempts = 5, baseMs = 250, maxMs = 5000,
  retryOn = [/Rate Limit/i, /Too many times/i, /unavailable/i, /429/, /5\d\d/]
} = {}) { /* Utilities.sleep(backoffWithJitter); try/catch fn(); */ }
```

---

## Suggested File Structure

```
lib/
├── core/
│   ├── Constants.js           # REQUIRED_CONFIG_KEYS, REQUIRED_PROJECT_COLUMNS, AUTOMATION_STATUS
│   ├── Utilities.js           # withBackoff(), formatDate(), parseReminderOffsets(), pure helpers
│   ├── Validator.js           # Validator class
│   └── ExecutionContext.js    # ExecutionContext class (wires everything together)
│
├── data/
│   ├── Config.js              # Config class
│   ├── IdAllocator.js         # IdAllocator class
│   ├── Project.js             # Project class (row-level wrapper)
│   ├── ProjectSheet.js        # ProjectSheet class (manages Project instances)
│   ├── SnapshotSheet.js       # SnapshotSheet class
│   ├── Directory.js           # Directory class
│   └── Codes.js               # Codes class
│
├── services/
│   ├── NotificationService.js # NotificationService class (includes internal template caching)
│   ├── ProjectService.js      # ProjectService class
│   ├── MaintenanceService.js  # MaintenanceService class
│   └── FormService.js         # FormService class
│
├── Main.js                    # Public API: processNewProjects(), runDailyMaintenance(), etc.
└── appsscript.json            # Manifest

client/
├── Client.js                  # Per-district thin client (SPREADSHEET_ID + trigger handlers)
└── appsscript.json            # Manifest (references library)
```

### Notes

- **Subdirectories for local navigation** — `core/`, `data/`, `services/` make it easy to find files in your IDE.
- **Apps Script flattening** — Clasp pushes these as `core/Constants`, `data/Config`, etc. in the Apps Script editor. Looks a bit ugly there, but you'll rarely edit in the browser anyway.
- **Load order** — Apps Script loads alphabetically by full path. `core/` loads before `data/` loads before `services/`, which naturally handles dependencies (constants → data classes → services).
- **`Main.js` at root** — Keeps the public API visually separate and easy to find.
- **`client/` is bound to the Main Projects File spreadsheet** — Each district gets their own copy. The script handles all triggers (form submit, edit, time-driven, manual menu). The manifest references the library by script ID. No separate form-bound script is needed.

