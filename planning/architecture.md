### 1\. Sheets & structures

* **Main Projects File**  
  * Row 1: user-facing column labels.  
  * Row 2: internal keys (e.g. `project_id`, `project_name`, `project_status`, `automation_status`, `due_date`, `project_lead_email`, `assignee_emails`, `folder_id`, `calendar_event_id`, `reminder_offsets`, etc.).  
  * Rows 3+: one project per row.  
* **Staff Directory**  
  * Columns: `email`, `name`, `role`, `active`, etc.  
  * Used for validating assignee/lead emails and powering dropdowns.  
* **Status Snapshot sheet**  
  * One row per project: `project_id`, `project_status` (and any other fields we want to compare daily).  
  * Overwritten each day after the daily digest runs.  
* **Reminder Profiles / Offsets sheet**  
  * Stores default reminder offsets and labels (e.g. `3`, `7`, `14` days before).  
  * Used to power dropdown choices and documentation.  
* **Templates config sheet**  
  * Maps project category → template file IDs (docs, sheets, etc.).

---

### 2\. Column indexing strategy

* On script startup:  
  * Read row 2 of Main Projects File.  
  * Build `key → column index` map.  
  * Validate:  
    * All required keys in `EXPECTED_KEYS` exist and are unique.  
  * If invalid, throw error and/or email admin.  
* All subsequent reads/writes use keys, not hard-coded column numbers.

---

### 3\. Project ID generation

* Format: `<DIST>-yy_yy-####` (e.g. `NUSD-25_26-0024`).  
* Store `(district, year, next_serial)` in a config sheet and update it with ScriptLock.  
* The Project ID is written to `project_id` column.

---

### 4\. Automation Status lifecycle

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

---

### 5\. Triggers and flows

* **Form submission trigger**  
  * Normalizes form input into a Main Projects File row.  
  * Sets `automation_status` to `Ready`. (MAYBE) manually kicks off batch trigger run.  
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
      * Send email to responsible people and invite them to the calendar event. We will have a Google Doc email template.  
  * If Automation Status is set to `Updated`, the 10-minute automation will re-sync the calendar event's date, attendees, and details with the current values in the Main Projects File, send the corresponding update notifications to the project lead and assignees, and then set Automation Status back to `Created`.  
  * If Automation Status is set to `Delete (Notify)` or `Delete (Don't Notify)`, the 10-minute automation will cancel the linked calendar event using the stored `calendar_event_id`, optionally send a cancellation notice to attendees based on the chosen option, hide the project row for archival purposes, and then set Automation Status to `Deleted`.  
  * Note: The `Updated` status handles explicit user-requested changes. The daily calendar sync (below) serves as a safety net to catch any discrepancies that may have been missed.  
* **Manual “Run now”**  
  * Custom menu or button calling `processNewProjects()`.  
* **Time-driven daily trigger (e.g. 8am)**  
  * Function: `dailyMaintenance()`.  
  * Responsibilities:  
    1. **Reminders**:  
       * For each active project, parse `reminder_offsets` (e.g. `"3,7,14"`).  
       * For each offset, compute `due_date - offset`; if equals today, send reminder email to assignees.  
    2. **Status-change digest**:  
       * Join current project statuses to Status Snapshot by `project_id`.  
       * For rows where `project_status` changed, group by person (lead \+ assignees).  
       * Send one email per person summarizing changed projects.  
       * Overwrite Status Snapshot with latest statuses.  
    3. **Late status**:  
       * For projects where due date is today and status is not “Completed”, set `project_status = Late` (or similar).  
    4. **Calendar sync (safety net)**:  
      * For each project:  
        * If due date ≠ calendar event date → update event date and send notification.  
        * If lead/assignee emails differ from event attendees → update attendees and send notification.  
      * This serves as a backup to catch any discrepancies that the `Updated` status workflow may have missed.

---

### 6\. Calendar/event data

* Use a single “robo” calendar per district for all project events.  
* Store `calendar_event_id` in Main Projects File for each project.  
* Event details:  
  * Title: `[Project ID] Project Name`.  
  * Description: link to project folder and summary fields.  
* All updates (date, attendees, deletion) are done by referencing `calendar_event_id` rather than trying to re-find events by title.

---

### 7\. Reminder offsets implementation

* Reminder offsets in Main Projects File:  
  * Multi-select dropdown stores user-friendly values (e.g. `3 days before,1 week before,2 weeks before` representing days before due date).
  * Human-friendly labels (e.g. "3 days before", "1 week before") live in helper columns / the Reminder Profiles/Offsets sheet - they have mappings built in to integer values, the users can edit integer values and then formulas auto-shift the dropdown options. This allows users to easily change the configuration.

---

### 8\. Folder \+ templates

* On `Created`:  
  * Create folder named `"Project Name [Project ID]"` under a district-level parent folder.  
  * Copy relevant templates based on project category.  
  * Run token substitution in templates (project name, due date, lead name, etc.).  
* Store `folder_id` in the row and use it for links in emails and calendar event descriptions.
