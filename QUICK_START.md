# i2i Teaming Tool - Quick Start Guide

## Key Features

- **Automated Project Setup**: Creates Google Drive folders, calendar events, and sends notifications when projects are submitted
- **Dual Submission Methods**: Submit via Google Form or manually enter in the spreadsheet
- **Smart Status Tracking**: Tracks project status changes and sends daily digest emails
- **Automated Reminders**: Configurable deadline reminders (e.g., 3 days, 1 week, 2 weeks before due)
- **Calendar Sync**: Automatically syncs calendar events when deadlines or assignees change
- **Staff Directory Integration**: Form dropdowns automatically update when staff are added/removed

---

## Core Components

**Main Projects File** (Google Sheets)
- Central database for all projects
- Two header rows: Row 1 = user-facing labels, Row 2 = internal keys (don't modify)
- Contains project details, status, automation status, and hidden system IDs

**Project Submission Form** (Google Forms)
- Primary method for submitting new projects
- Dropdowns automatically sync with Staff Directory and Category codes
- Form responses are processed immediately

**Staff Directory** (Sheet within Main Projects File)
- Manages team members: Name, Email, Permissions (Edit/View/No Access)
- Changes automatically update form dropdowns
- Permissions control access to the Main Projects File

**Config Sheet** (Hidden sheet)
- Stores district-specific settings: District ID, School Year, folder IDs, template IDs
- Required for automation to function

**Status Snapshot Sheet** (Hidden sheet)
- Tracks previous day's project statuses for change detection
- Used by daily maintenance to generate status change digests

---

## Submitting Projects

### Via Google Form
1. Fill out the Project Submission Form
2. Submit — automation processes immediately
3. System sets `Automation Status` to `Ready`, then processes to `Created`

### Manual Entry
1. Add a new row in the Main Projects File
2. Fill in required fields: Title, Description, Assigned to, Requested by, Deadline, Category
3. Set `Automation Status` dropdown to `Ready` when the row is complete
4. Background automation (runs every 10 minutes) picks it up and processes it

### Automation Status Workflow

The `Automation Status` column controls when and how projects are processed. The dropdown options change based on the current state:

**Blank → Ready**
- For new manual entries: Set to `Ready` when all required fields are filled
- For form submissions: Automatically set to `Ready` by the system

**Ready → Created**
- Background automation (every 10 minutes) processes `Ready` projects:
  - Creates project folder in Google Drive
  - Copies template file into folder
  - Creates calendar event with deadline
  - Sends notification email to assignees
  - Sets status to `Created` on success, or `Error` if something fails

**Created → Updated**
- When project details change (deadline, assignees, etc.):
  - Set `Automation Status` to `Updated`
  - Background automation re-syncs calendar event and sends update notifications
  - Status returns to `Created` after sync completes

**Created → Delete**
- To cancel a project:
  - Set to `Delete (Notify)` to cancel calendar event and email attendees
  - Set to `Delete (Don't Notify)` to cancel event without email
  - Background automation cancels event, hides the row, sets status to `Deleted`

**Error → Ready**
- If processing fails, status becomes `Error`
- Fix the issue (missing folder ID, invalid email, etc.)
- Set back to `Ready` to retry

**Important**: The dropdown validation enforces these transitions. You can only select valid next states based on the current status.

---

## Managing Projects

**Viewing Projects**
- All active projects appear in the Main Projects File
- Columns include: Project ID, Title, Category, Assigned to, Deadline, Project Status, Automation Status
- Hidden columns store system IDs (folder_id, calendar_event_id) — don't edit these

**Project IDs**
- Format: `DIST-yy_yy-####` (e.g., `NUSD-25_26-0024`)
- Auto-generated based on District ID and School Year from Config sheet
- Serial number increments automatically

**Updating Project Status**
- Use the `Project Status` dropdown to track progress
- Default options: Not started, Behind schedule, Stuck, On track, Complete, Late
- Status changes are tracked daily for digest emails
- Setting status to "Complete" automatically records a `Completed At` timestamp

**Changing Project Details**
- Edit deadline, assignees, or other fields directly in the sheet
- Set `Automation Status` to `Updated` to sync changes to calendar event
- Background automation processes updates within 10 minutes

---

## Automated Features

**On Project Submission** (Form or Manual with `Ready` status)
- Project folder created in designated parent folder
- Template file copied into project folder (with project details auto-filled)
- Calendar event created with deadline
- Notification email sent to assignees (cc: requester)
- Project ID assigned and recorded

**Daily Maintenance** (Runs at 8am)
- **Reminder Emails**: Sends deadline reminders based on each project's reminder timeline (e.g., 3 days, 7 days, 14 days before due)
- **Status Change Digest**: Compares current statuses to yesterday's snapshot, sends digest email to project leads and assignees listing all status changes
- **Auto-Late Status**: Sets `Project Status` to "Late" on the day a project is due (if not already Complete)
- **Calendar Sync**: Checks if deadline or assignees changed, updates calendar event and sends notifications

**Background Processing** (Every 10 minutes)
- Processes projects with `Automation Status` = `Ready`, `Updated`, or `Delete`
- Creates folders/events for new projects
- Re-syncs calendar events for updated projects
- Cancels events and hides rows for deleted projects

**Manual Triggers** (Teaming Tool menu)
- **Run Now**: Immediately process all `Ready`/`Updated`/`Delete` projects (don't wait for 10-minute cycle)
- **Sync Form Dropdowns**: Refresh form dropdowns with latest Directory and Category codes
- **Refresh Permissions**: Re-apply sharing settings based on Staff Directory permissions

---

## Notifications & Communication

**New Project Assignment**
- Sent to assignees when project is created
- Includes: Project title, category, deadline, description, folder link
- CC: Person who requested the project

**Reminder Notifications**
- Sent to assignees only (not requester)
- Triggered based on `Reminder Timeline` setting per project
- Default timeline: 3 days, 7 days, 14 days before deadline
- Includes days remaining and folder link

**Status Change Digest**
- Daily email (8am) to project leads and assignees
- Lists all projects where status changed in the past 24 hours
- One email per person (digest format, not per-project)

**Project Update Notification**
- Sent when `Automation Status` is set to `Updated` and changes are synced
- Includes summary of what changed (deadline, team members, etc.)
- Sent to all project participants

**Project Cancellation**
- Sent when `Automation Status` is set to `Delete (Notify)`
- Notifies attendees that calendar event is cancelled
- Includes project details and cancellation reason

---

## Staff Management

**Adding Staff**
1. Add row to Directory sheet: Name, Email Address, Permissions
2. Permissions: `edit` (or `editor`) = can edit Main Projects File, `view` (or `viewer`) = read-only, blank = no access
3. Form dropdowns automatically update within a few minutes (or use "Sync Form Dropdowns" menu)

**Removing Staff**
1. Delete row from Directory sheet
2. Run "Refresh Permissions" from menu to remove their access to Main Projects File
3. Form dropdowns automatically update

**Updating Permissions**
1. Change Permissions value in Directory sheet
2. Run "Refresh Permissions" from menu to apply changes
3. System adds/removes/upgrades permissions as needed

**Important**: Directory changes don't automatically update file permissions — use the "Refresh Permissions" menu option to sync.

---

## Project Folders & Templates

**Folder Creation**
- Automatically created when project status is `Ready` → `Created`
- Located in the parent folder specified in Config sheet
- Folder name: Project ID + Project Title
- Shared with assignees and requester

**Template File**
- Template Google Sheets file (specified in Config) is copied into each project folder
- Template's Overview tab is auto-filled with project details (name, deadline, assignee, etc.)
- Template structure is defined by the district — customize as needed

**Accessing Folders**
- Folder link is included in all notification emails
- Folder ID is stored in hidden column (for system reference)
- Folders remain accessible even if project is deleted (row is hidden, folder is not deleted)

---

## Troubleshooting & Manual Actions

**"Run Now" Menu Option**
- Use when you want immediate processing instead of waiting for the 10-minute cycle
- Processes all `Ready`, `Updated`, and `Delete` projects
- Helpful after bulk manual entries or when testing

**"Sync Form Dropdowns" Menu Option**
- Refreshes form dropdowns with latest Directory and Category codes
- Use after adding/removing staff or changing category options
- Also runs automatically when Directory or Codes sheets are edited

**"Refresh Permissions" Menu Option**
- Re-applies file sharing based on Staff Directory permissions
- Use after updating Directory permissions
- Adds missing permissions, removes people not in Directory, upgrades/downgrades as needed

**Common Issues**

**Project stuck at `Ready` status**
- Check for errors in required fields (invalid email, missing folder ID in Config, etc.)
- Check Apps Script execution logs for error messages
- If status shows `Error`, fix the issue and set back to `Ready`

**Calendar event not updating**
- Ensure `Automation Status` is set to `Updated` (not just editing the deadline field)
- Wait for 10-minute cycle or use "Run Now"
- Check that calendar event ID is present in hidden column

**Form dropdowns not updating**
- Use "Sync Form Dropdowns" menu option
- Verify Directory sheet has correct column headers: `Name`, `Email Address`, `Permissions`
- Check that staff emails are valid

**Notifications not sending**
- Verify email template Google Doc IDs in Config sheet
- Check that assignee emails are valid and in Directory
- Review Apps Script execution logs for email errors

