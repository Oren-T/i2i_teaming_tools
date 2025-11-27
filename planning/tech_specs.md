#### Project submission form

- A form to submit new projects  
- Likely to use Google Forms, not something more custom-built, since Google Forms has easier permissions management. (Joel to either create the form, or send Oren list of form questions to include.)  
  - Note from the feedback session: there will be some kind of category list in here.  
- **Staff directory** is linked to **Project Submission Form** \- when new staff are added/removed from a directory file, their names automatically appear / disappear from form dropdowns

#### Manual project submissions

- Users can manually enter projects directly into the **Main Projects File**.  
- There is a column called Automation Status. For manual entries, users set this to Ready once they’ve finished entering all required details.  
  1. This allows them to wait until they are finished entering all details before turning it to Ready.  
  2. The field can be: blank, Ready, Created  
- (When a new project is submitted through the form, we set it to Ready and then the code flips it to Created once all of the programmed automation is complete.)  
- A background automation checks for new Ready projects approximately every 10 minutes (or when a user manually runs it) and creates all associated assets (project folder, templates, calendar event, etc.). Once this is complete, Automation Status is set to Created.

#### Staff directory

- A file to house staff names, emails, and permissions levels.  
  - We decided that for now, each district will only receive one instance of the Teaming Tool. Thus, permission levels will simply toggle edit/view/no access to the main projects file overall. (There are not multiple files to juggle permissions for.)  
- When changes are made to this, we can implement automations to alter any access, etc. The most critical implementation will be to ensure new staff names properly appear in the **Project submission form**.

#### Main Projects File

- One record per project  
- Project status is recorded as a change-able column in this sheet  
  * We will let people use a dynamic dropdown list of task statuses \- the suggested initial configuration will be: Not started, behind schedule, stuck, on track, completed  
  * Project statuses can be manually updated by team members who have permission to access this sheet. Status changes are tracked via the **Status Snapshot sheet** (see below), which compares current statuses to the previous day's snapshot to detect changes. Other columns in this sheet should be protected from most peoples' editing abilities.  
- Automatic updates to the **Main Projects File** are triggered by one of two events:  
  1. A submission is made to the **Project submission form**  
  2. Manual entry of a new project row in the Main Projects File, when the Automation Status is set to Ready \- a background automation runs roughly every 10 minutes (and can be manually triggered from within the sheet) to see if any new projects have been manually created.  
- When either of these events occur, the following actions will be triggered:  
  1. A **Project Folder** will be automatically created (details below)  
  2. An email notification will be sent to the project assignee(s), copying the leader who submitted the project, letting them know a new project has been submitted and assigned to them. This includes the link to the **Project Folder**, and basic information about the project and deadline.  
  3. A calendar invite is sent to project assignees and the leader who submitted the project.  
     - This uses a dedicated "robo" Google account within the school server to host the calendar events.  
  4. The ID of each project folder and event is stored in a hidden column in the **Main Projects File** so that calendar events and notifications can link back to the same folder, even if names or locations change.  
- Every day, a script will run which…  
  1. Determines if reminder notifications need to be sent (to the assignee only) for any currently incomplete projects.  
     - Sends email notifications to assigned staff of upcoming project deadlines at a pre-defined cadence. There will be a default cadence (e.g. 3 days before, 1 week before, 2 weeks before), but this can be changed on a per-project level using dropdown menus in this spreadsheet.
     - The dropdown menu will display user-friendly labels (e.g. "3 days before"), which map to integer values stored in helper columns/sheets. This allows for easy configuration changes.  
  2. The project lead and all project assignees should receive an email notification if any project they are assigned to changed statuses in the past 24 hours. This is detected by comparing current project statuses to the **Status Snapshot sheet** (see below). It probably makes sense to do this as a per-person daily digest \- i.e. "here are the projects you are on where the status changed in the past 24 hours," to avoid a single person getting many notifications per day.  
  3. Auto-sets project status to Late on the day a project is due  
  4. Checks to see if a project due date in this spreadsheet is different from the date of the calendar event (i.e. if the deadline changed). If the deadline changed, the calendar event date should be shifted, and an appropriate notification should be sent.  
  5. Additionally, check to see who is listed as a responsible person for the project and who else is assigned to it. Compare this list to the live calendar event. If this list has changed since the calendar event was created, make appropriate changes to the calendar event and send out an associated email notification.  
- Headers  
  1. The Main Projects File will use two header rows:  
     - Row 1: user-facing labels (can be renamed/reordered by admins).  
     - Row 2: internal “key” names that the automation depends on (e.g. project\_id, due\_date, project\_status, automation\_status). These should not be changed except by the tool maintainer.  
- IDs & hidden columns  
  1. Each project will have a unique Project ID of the form DIST-yy\_yy-\#\#\#\# (e.g. NUSD-25\_26-0024), where \#\#\#\# is the project’s sequence number for that district and school year.
     - The academic year (yy_yy) and the current serial number are stored in and retrieved from a config sheet in the Main Projects spreadsheet.  
  2. The row will also store hidden columns for:  
     - Project folder ID (Google Drive)  
     - Calendar event ID  
     - Any other system IDs needed for synchronization.  
- Automation Status (creation \+ deletion)  
  1. The sheet includes an Automation Status column, distinct from the user-visible Project Status.  
  2. Typical Automation Status flow:  
     - Blank → Ready → Created → (optionally) Delete (Notify) / Delete (Don’t Notify) / (optionally) Updated.  
  3. Users set Ready when a new row is complete. The tool marks the project as Created once all automated setup is finished. After that:  
     - Users can request deletion of the calendar event via Delete (Notify) or Delete (Don’t Notify).  
     - Users can signal an update to project details via Updated.  
     - A background script (same one every 10 minutes) processes these requests:
       - For Delete: cancels the event and hides (but does not delete) the row.
       - For Updated: re-syncs the calendar event date/attendees/details and sends notifications.

#### Status Snapshot sheet

- A separate sheet within the Main Projects spreadsheet that tracks project statuses for comparison purposes.  
- Structure:  
  * Row 1: Header row with column labels: `project_id`, `project_status`  
  * Rows 2+: One row per project, storing the `project_id` and `project_status` from the previous day  
- Purpose:  
  * Used by the daily maintenance script to detect which projects have changed status since the last run  
  * The script compares current `project_status` values from the Main Projects File to the snapshot  
  * After sending status-change digest emails, the snapshot is overwritten with the current statuses  
- Initialization:  
  * On first run, if the snapshot is empty or missing projects, the script should populate it with all current projects  
  * New projects added to the Main Projects File should be added to the snapshot on their first daily run (with their current status)

#### Project Folder

- When a project folder gets created, a single pre-generated template file (Google Sheets) will be copied into the project folder.  
- Joel will create the template file.  
- Custom tokens can be auto-substituted into the template's Overview tab \- e.g. Project Name, deadline, assignee, etc. are auto-filled during the copy process.  
- In the old system, Project Status lived within a **Project Folder** google sheet for tracking the project. However, now, we have moved Project Status to live in the **Main Projects File** instead, and be manually updated by users.
