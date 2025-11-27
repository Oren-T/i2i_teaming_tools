### 1\. Sheets & structures

* **Main Projects File**  
  * Row 1: user-facing column labels.  
  * Row 2: internal keys (e.g. `project_id`, `project_name`, `project_status`, `automation_status`, `due_date`, `requested_by`, `assignee`, `folder_id`, `calendar_event_id`, `reminder_offsets`, etc.).  
  * Rows 3+: one project per row.  
* **Staff Directory**  
  * Columns: `Name`, `Email Address`, `Permissions`.  
  * Used for validating assignee/lead emails and powering dropdowns.  
  * Active status is inferred from `Permissions` column (empty/none/no access = inactive).  
* **Status Snapshot sheet**  
  * Structure:  
    * Row 1: Header row with column labels: `project_id`, `project_status`  
    * Rows 2+: One row per project, storing `project_id` and `project_status` from the previous day  
  * Purpose: Tracks previous day's project statuses to detect changes during daily maintenance  
  * Lifecycle:  
    * On first run: Populate snapshot with all current projects from Main Projects File  
    * Daily: Compare current statuses to snapshot, identify changes, send digest emails, then overwrite snapshot with current statuses  
    * New projects: Added to snapshot on their first daily run (with their current status at that time)  
* **Reminder Profiles / Offsets sheet**  
  * Stores default reminder offsets and labels (e.g. `3`, `7`, `14` days before).  
  * Used to power dropdown choices and documentation.  

---

### 2\. Deployment Strategy (Library vs Client)

* **Architecture**: "Thin Client" script + Central Library.
* **Central Library**:
  * Contains all business logic (`processNewProjects`, `dailyMaintenance`, `generateId`).
  * Managed in a central repo/script project.
  * Versioned updates.
* **Client Script (per District)**:
  * Minimal code; acts as a bridge.
  * Contains `CONFIG` object (District ID, Year, etc - though some of this may end up living in a user-facing sheet).
  * Defines simple triggers (`onOpen`, `timeDriven`) that strictly pass execution to the Library functions with the local config.
  * **Benefit**: Updates can be pushed to all districts by updating the library version, without editing individual district scripts.

---

### 3\. Column indexing strategy

* On script startup:  
  * Read row 2 of Main Projects File.  
  * Build `key → column index` map.  
  * Validate:  
    * All required keys in `EXPECTED_KEYS` exist and are unique.  
  * If invalid, throw error and/or email admin.  
* All subsequent reads/writes use keys, not hard-coded column numbers.

---

### 4\. Project ID generation

* Format: `<DIST>-yy_yy-####` (e.g. `NUSD-25_26-0024`).  
* Store `(district, year, next_serial)` in a config sheet and update it with ScriptLock.  
* The Project ID is written to `project_id` column.

---

### 5\. Automation Status lifecycle

* `automation_status` values and transitions:  
  * `''` → user fills row.  
  * `Ready` → user signals row is complete (manual entry) or automation sets it after form submission normalization. There must be a script lock to ensure manual trigger doesn’t conflict with timed automation.  
  * Batch job picks up `Ready` rows:  
    * Creates folder \+ templates.  
    * Creates calendar event.  
    * Writes `folder_id` and `calendar_event_id`.  
    * Sets `automation_status` → `Created` on success (or `Error` if something fails).  
  * Error handling: If status is `Error`, users can manually set it back to `Ready` to retry the batch job processing.  
  * Later, user can set:  
    * `Delete (Notify)` or `Delete (Don't Notify)` to request calendar deletion; or, `Updated` to set a notification that the information has been updated. The script running every 10 minutes should detect updates and respond accordingly, then change the status back to `Created` once the update has been processed.  
* Enforcement: drop-down validation. For a blank row, the only option in the dropdown is `Ready` (reserve `Created` for the script). Once the code changes a status to `Created`, the dropdown options change to be only `Created`, `Updated`, `Delete (Notify)`, or `Delete (Don't Notify)`. If status is `Error`, users can set it back to `Ready` to retry.

#### State Machine

| Current Status | Valid Next States | Triggered By | Action |
|----------------|-------------------|--------------|--------|
| (blank) | Ready | User | User finished entering row |
| Ready | Created, Error | Automation | Batch processes the row |
| Created | Updated, Delete (Notify), Delete (Don't Notify) | User | User requests change |
| Updated | Created | Automation | Re-sync complete |
| Delete (Notify) | Deleted | Automation | Cancel event, notify attendees, hide row |
| Delete (Don't Notify) | Deleted | Automation | Cancel event, hide row |
| Error | Ready | User | User retries after fixing issue |
| Deleted | (terminal) | — | Row is hidden, no further processing |

#### Dropdown Validation Rules

| Row State | Dropdown Options Available |
|-----------|---------------------------|
| Blank row (new entry) | `Ready` only |
| `Created` | `Created`, `Updated`, `Delete (Notify)`, `Delete (Don't Notify)` |
| `Error` | `Ready` (to retry) |
| Other states (`Ready`, `Updated`, `Delete *`, `Deleted`) | Locked (no user edits) |

---

### 6\. Triggers and flows

* **Form submission trigger**  
  * **Architecture decision:** The Google Form is linked to send responses to a **hidden "Form Responses (Raw)" tab** within the Main Projects File spreadsheet. This allows a single spreadsheet-bound script to handle all triggers (no separate form-bound script needed).
  * **Ingestion Flow**: Google Form -> Hidden "Form Responses (Raw)" tab -> `onFormSubmit` trigger -> Script normalizes and appends to Main Projects sheet.  
  * **Process**:  
    * `onFormSubmit` trigger (installable, on the spreadsheet) activates when a form response is submitted.
    * Script reads the raw response from the hidden tab, normalizes the data (mapping form questions to project keys), and appends it to the main **Projects** sheet.  
    * Sets `requested_by` from the form submitter's email address (automatically captured by Google Forms, then looked up in Directory for the name).  
    * Sets `automation_status` to `Ready`.  
    * (Optional) Immediately kicks off project processing rather than waiting for the 10-minute batch.
* **Time-driven batch trigger (every 10 minutes)**  
  * Function: `processNewProjects()`.  
  * Acquires ScriptLock at the start to prevent overlapping runs (manual trigger vs timed trigger) from processing the same `Ready`/`Updated`/`Delete` rows simultaneously.  
  * Steps:  
    * Read Main Projects File into memory.  
    * Filter rows with `automation_status = Ready`.  
    * For each:  
      * Generate Project ID if missing. If `project_id` already exists but status is `Ready`, log/email the error for that row, set status to `Error`, skip it, and continue processing the rest.  
      * Create project folder \+ copy templates, run token substitution.  
      * Create calendar event on district “robo” calendar.  
      * Write `folder_id`, `calendar_event_id`, `project_id` back to row.  
      * Set `automation_status = Created`.  
      * Send email to responsible people and invite them to the calendar event. Email templates are stored as separate Google Docs (one per template type). First line = subject, remaining lines = body. Script parses by splitting on newlines and performs token substitution.  
  * If Automation Status is set to `Updated`, the 10-minute automation will re-sync the calendar event's date, attendees, and details with the current values in the Main Projects File, send the corresponding update notifications to the project lead and assignees, and then set Automation Status back to `Created`.  
  * If Automation Status is set to `Delete (Notify)` or `Delete (Don't Notify)`, the 10-minute automation will cancel the linked calendar event using the stored `calendar_event_id`, optionally send a cancellation notice to attendees based on the chosen option, hide the project row for archival purposes, and then set Automation Status to `Deleted`.  
  * Note: The `Updated` status handles explicit user-requested changes. The daily calendar sync (below) serves as a safety net to catch any discrepancies that may have been missed.  
* **Manual “Run now”**  
  * Custom menu or button calling `processNewProjects()`.  
* **Permissions Management (Manual)**
  * Custom menu item: "Refresh Permissions".  
  * Action: Re-applies sharing settings to the Main Projects File based on the Staff Directory roles. Not strictly enforced by a timer, but available to fix discrepancies.
* **Form Dropdown Sync (Manual)**
  * Custom menu item: "Refresh Form Dropdowns".  
  * Action: Updates Google Form dropdown options (Category and Assigned to) from Codes and Directory sheets. Available for manual refresh if auto-sync misses updates.
* **Main Projects File onEdit trigger**
  * Trigger: `onEdit` on the Main Projects File spreadsheet (handles edits to Main Projects sheet, Staff Directory sheet, and Codes sheet).
  * Actions:
    * If `project_status` column edited and new value is "Completed" (and old value was not), set `completed_at` to current timestamp.
    * If Staff Directory sheet edited, update Google Form dropdown options (Assigned to) to match currently active staff.
    * If Codes sheet Category column edited, update Google Form dropdown options (Category) to match current categories.
* **Time-driven daily trigger (e.g. 8am)**  
  * Function: `dailyMaintenance()`.  
  * Responsibilities:  
    1. **Reminders**:  
       * For each active project, parse `reminder_offsets` (e.g. `"3,7,14"`).  
       * For each offset, compute `due_date - offset`; if equals today, send reminder email to assignees.  
    2. **Status-change digest**:  
       * Read current project statuses from Main Projects File (all rows with `project_id` and `project_status`).  
       * Read previous statuses from Status Snapshot sheet.  
       * Join by `project_id` and identify rows where `project_status` differs between current and snapshot.  
       * For projects with status changes:  
         * Group by person (requested_by + assignee).  
         * Send one email per person summarizing all projects where status changed.  
       * After sending all digest emails, overwrite Status Snapshot with current statuses (all active projects).  
       * Note: If a project exists in Main Projects File but not in snapshot, add it to snapshot. If a project exists in snapshot but not in Main Projects File (deleted/hidden), remove it from snapshot.  
    3. **Late status**:  
       * For projects where due date is today and status is not “Completed”, set `project_status = Late` (or similar).  
    4. **Calendar sync (safety net)**:  
      * For each project:  
        * If due date ≠ calendar event date → update event date and send notification.  
        * If lead/assignee emails differ from event attendees → update attendees and send notification.  
      * This serves as a backup to catch any discrepancies that the `Updated` status workflow may have missed.

---

### 7\. Calendar/event data

* Use a single "robo" calendar per district for all project events.  
* **Implementation**: Script runs as a bot account (e.g., `calendar-bot@yourdomain.com`). Uses `CalendarApp.getDefaultCalendar()` to access the bot's primary calendar. No calendar ID needed in config.  
* Store `calendar_event_id` in Main Projects File for each project.  
* Event details:  
  * Title: `[Project ID] Project Name`.  
  * Description: link to project folder and summary fields.  
* All updates (date, attendees, deletion) are done by referencing `calendar_event_id` rather than trying to re-find events by title.

---

### 8\. Reminder offsets implementation

* Reminder offsets in Main Projects File:  
  * Multi-select dropdown stores user-friendly values (e.g. `3 days before,1 week before,2 weeks before` representing days before due date).
  * Human-friendly labels (e.g. "3 days before", "1 week before") live in helper columns / the Reminder Profiles/Offsets sheet - they have mappings built in to integer values, the users can edit integer values and then formulas auto-shift the dropdown options. This allows users to easily change the configuration.

---

### 9\. Folder \+ templates

* On `Created`:  
  * Create folder named `"Project Name [Project ID]"` under a district-level parent folder.  
  * Copy relevant templates based on project category.  
  * Run token substitution in templates (project name, due date, lead name, etc.).  
* Store `folder_id` in the row and use it for links in emails and calendar event descriptions.
