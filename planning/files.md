### Main Projects File (Google Sheets)

* **Project Management Sheet** - Main data sheet with all project records (two header rows: user labels + internal keys)
  * Row 1 (user-facing labels): Project ID, Created At, School Year, Goal #, Action #, Category (default is LCAP), Title, Description, Assigned to, Requested by, Deadline, Project Status, Completed At?, Reminder Timeline, Automation Status, Calendar Event ID, Folder ID, Notes
  * Row 2 (internal keys): `project_id`, `created_at`, `school_year`, `goal_number`, `action_number`, `category`, `project_name`, `description`, `assignee`, `requested_by`, `due_date`, `project_status`, `completed_at`, `reminder_offsets`, `automation_status`, `calendar_event_id`, `folder_id`, `notes`

* **Status Snapshot** - Hidden sheet tracking previous day's project statuses for change detection (`project_id`, `project_status`)

* **Codes** - Dropdown content: Category, Status, Reminder Days Offset, Reminder Days: Readable. Updates here change dropdowns on main page.
  * **Category options:** LCAP, SPSA, Community School, WASC, Other (default: LCAP)

* **Directory** - Staff directory. Column headers: `Name`, `Email Address`, `Permissions`. Powers form dropdowns.

* **Config** - System configuration (see details below)

**Config Sheet Structure:**

*Stores district-specific configuration needed for automation. Script uses header cell constants (e.g., `Config!A2` for "District ID") rather than header-based lookup because the structure is stable and rarely changed; `getDataRegion()` can be used from there.*

| key | value | description |
|-----|-------|-------------|
| District ID | NUSD | District abbreviation for Project ID format (e.g., NUSD-25_26-0024) |
| School Year | 25_26 | Current school year in yy_yy format |
| Next Serial | 1 | Next project serial number (auto-incremented with ScriptLock) |
| Parent Folder ID | | Google Drive folder ID where project folders are created |
| Project Template ID | | Google Sheets file ID of the Project File Template to copy |
| Form ID | | Google Form ID for Project Submission Form (for form sync) |
| Error Email Addresses | | Comma-separated list of email addresses to notify on automation errors |
| Email Template - New Project | | Google Doc ID for new project assignment email template |
| Email Template - Reminder | | Google Doc ID for reminder email template |
| Email Template - Status Change | | Google Doc ID for status change digest email template |
| Email Template - Project Update | | Google Doc ID for project update notification email template |
| Email Template - Project Cancellation | | Google Doc ID for project cancellation notification email template |

---

### Project File Template (Google Sheets)

*Single template copied into each project folder. Tabs: Overview, Action Plan, Meeting Notes, Progress Monitoring, Resources and Glossary, Codes.*

**Overview Tab Structure:**
* Column A (A2:A10): Field labels
* Column B (B2:B10): Associated values (populated from Main Projects File during template copy)

| Row | Label (Column A) | Value Source (Column B) |
|-----|------------------|------------------------|
| 2 | School Year | |
| 3 | Goal # | |
| 4 | Action # | |
| 5 | Category (default is LCAP) | |
| 6 | Title | |
| 7 | Description | |
| 8 | Assigned to | |
| 9 | Requested by | |
| 10 | Deadline | |

---

### i2i Teaming Tool Library (Google Apps Script)

*Central library with all business logic, shared across district instances*

---

### Project Submission Form (Google Form)

*Form submissions flow to Main Projects File via form submission trigger. Dropdowns populated from Directory.*

**Form Fields:**
1. Goal # - Short answer text
2. Action # - Short answer text
3. Category - Dropdown from Codes sheet (default: LCAP)
4. Title - Required short answer text
5. Description - Required paragraph text
6. Assigned to - Multi-select dropdown from Directory
7. Deadline - Date picker
8. Notes - Optional paragraph text

---

### Email Templates (Google Docs)

*Three separate Google Docs, one per template type. File IDs stored in Config sheet. Token substitution used for project-specific details.*

**Structure:** First line = subject line, remaining lines = email body. Apps Script parses by splitting on newlines.

* **Email Template - New Project** - Sent to assignees when project is created, includes project details and folder link
* **Email Template - Reminder** - Sent to assignees at reminder intervals before deadline
* **Email Template - Status Change** - Sent in daily digest when project status changes
