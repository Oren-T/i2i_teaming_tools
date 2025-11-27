### Form Submission Trigger

*Event-driven / On form submit / Google Form*

* Reads form response, normalizes data, appends to Main Projects File, sets `automation_status` to `Ready`

---

### Time-Driven Batch Trigger (Every 10 Minutes)

*Time-driven / Every 10 minutes / Main Projects File*

* Processes rows with `automation_status = Ready/Updated/Delete`, creates folders/calendar events, handles deletions and updates

---

### Main Projects File onEdit Trigger

*Event-driven / On edit / Main Projects File spreadsheet (all tabs)*

* Detects `project_status` edits, sets `completed_at` timestamp when status becomes "Completed"
* Updates Google Form dropdown options when Staff Directory sheet is edited
* Updates Google Form dropdown options when Codes sheet (Category column) is edited

---

### Time-Driven Daily Trigger (8am)

*Time-driven / Daily at 8am / Main Projects File*

* Sends reminder emails, generates status-change digest, sets late status, syncs calendar events

---

### Manual Triggers

*Manual / User-initiated / Main Projects File*

* Manual "Run now" option to immediately process ready projects
* Re-apply sharing settings based on Staff Directory roles
* Refresh form dropdowns (updates Category and Assigned to options from Codes and Directory sheets)

