// ============================================================
//  CODE — Department Task Tracker
// ============================================================
//  Entry points: processTask(), sendReminders(), createTriggers()
//  Requires the Advanced Calendar Service to be enabled.
// ============================================================


// --------------- Trigger Setup (run once) ---------------

/**
 * Creates both time-driven triggers. Run once from the Apps Script editor.
 * Safe to re-run — removes existing triggers for these functions first.
 */
function createTriggers() {
  const managed = ['processTask', 'sendReminders'];

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (managed.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('processTask')
    .timeBased()
    .everyMinutes(5)
    .create();

  ScriptApp.newTrigger('sendReminders')
    .timeBased()
    .atHour(8)
    .nearMinute(0)
    .everyDays(1)
    .create();

  console.log('Triggers created: processTask (every 5 min), sendReminders (daily ~8:00 AM)');
}

// --------------- Startup Validation ---------------

/**
 * Validates that required tabs and columns exist in the spreadsheet.
 * Sends an error email and throws on any missing tab or column.
 * @returns {{ ss, projectsSheet, tasksSheet, directorySheet,
 *             projectColMap, taskColMap, directoryColMap }}
 */
function validateSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errors = [];

  const projectsSheet  = ss.getSheetByName(TAB_NAMES.PROJECTS);
  const tasksSheet     = ss.getSheetByName(TAB_NAMES.TASKS);
  const directorySheet = ss.getSheetByName(TAB_NAMES.DIRECTORY);

  if (!projectsSheet)  errors.push(`Missing tab: "${TAB_NAMES.PROJECTS}"`);
  if (!tasksSheet)      errors.push(`Missing tab: "${TAB_NAMES.TASKS}"`);
  if (!directorySheet) errors.push(`Missing tab: "${TAB_NAMES.DIRECTORY}"`);

  if (errors.length) {
    const msg = errors.join('\n');
    sendErrorEmail(msg);
    throw new Error(msg);
  }

  const projectColMap  = buildColumnMap(projectsSheet);
  const taskColMap     = buildColumnMap(tasksSheet);
  const directoryColMap = buildColumnMap(directorySheet);

  function checkCols(tabName, colMap, expected) {
    for (const key of Object.keys(expected)) {
      if (colMap[expected[key]] === undefined) {
        errors.push(`Missing column "${expected[key]}" in "${tabName}"`);
      }
    }
  }

  checkCols(TAB_NAMES.PROJECTS,  projectColMap,  PROJECT_COLUMNS);
  checkCols(TAB_NAMES.TASKS,     taskColMap,     TASK_COLUMNS);
  checkCols(TAB_NAMES.DIRECTORY, directoryColMap, DIRECTORY_COLUMNS);

  if (errors.length) {
    const msg = errors.join('\n');
    sendErrorEmail(msg);
    throw new Error(msg);
  }

  return {
    ss, projectsSheet, tasksSheet, directorySheet,
    projectColMap, taskColMap, directoryColMap
  };
}

// --------------- Task Processing (5-min trigger) ---------------

/**
 * Processes task rows whose "Update Reminders" checkbox is checked.
 * Creates or updates all-day calendar events via the Advanced Calendar Service
 * with sendUpdates: 'all' so Google sends native invitation/update emails.
 */
function processTask() {
  const env = validateSetup();
  const { tasksSheet, taskColMap } = env;
  const directory = loadDirectory(env.directorySheet, env.directoryColMap);
  const projectStatuses = loadProjectStatusMap(env.projectsSheet, env.projectColMap);

  const data = tasksSheet.getDataRange().getValues();

  const col = {
    checkbox:  taskColMap[TASK_COLUMNS.UPDATE_REMINDERS],
    taskName:  taskColMap[TASK_COLUMNS.TASK_NAME],
    project:   taskColMap[TASK_COLUMNS.PROJECT],
    assignee:  taskColMap[TASK_COLUMNS.ASSIGNEE],
    deadline:  taskColMap[TASK_COLUMNS.DEADLINE],
    status:    taskColMap[TASK_COLUMNS.STATUS],
    eventId:   taskColMap[TASK_COLUMNS.CALENDAR_EVENT_ID]
  };

  const checkedCount = data.reduce((n, row, i) => n + (i > 0 && row[col.checkbox] === true ? 1 : 0), 0);
  console.log(`processTask: starting — ${checkedCount} checked row(s) found.`);

  let created = 0;
  let updated = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[col.checkbox] !== true) continue;

    const sheetRow = i + 1;
    const checkboxRange = tasksSheet.getRange(sheetRow, col.checkbox + 1);

    try {
      const taskName    = String(row[col.taskName] || '').trim();
      const project     = String(row[col.project]  || '').trim();
      const assigneeStr = String(row[col.assignee] || '').trim();
      const deadline    = row[col.deadline];
      const existingId  = String(row[col.eventId]  || '').trim();

      // --- Required-field validation (skip row, leave checkbox checked) ---
      if (!taskName) {
        console.warn(`Row ${sheetRow}: Task Name is missing.`);
        setAutoNote(checkboxRange, 'Task Name is required.');
        continue;
      }
      if (!assigneeStr) {
        console.warn(`Row ${sheetRow}: Assignee is missing.`);
        setAutoNote(checkboxRange, 'Assignee is required.');
        continue;
      }
      if (!(deadline instanceof Date)) {
        console.warn(`Row ${sheetRow}: Deadline is missing or not a valid date.`);
        setAutoNote(checkboxRange, 'Deadline must be a valid date.');
        continue;
      }

      // --- Project-status check (uncheck and skip if inactive) ---
      if (project) {
        const projStatus = projectStatuses.get(project);
        if (projStatus === PROJECT_STATUS.COMPLETE || projStatus === PROJECT_STATUS.CANCELLED) {
          console.log(`Row ${sheetRow}: project "${project}" is ${projStatus} — unchecking.`);
          clearAutoNote(checkboxRange);
          checkboxRange.setValue(false);
          continue;
        }
      }

      // --- Resolve assignee names → emails ---
      const emails = resolveAssigneeEmails(assigneeStr, directory);
      if (emails.length === 0) {
        console.warn(`Row ${sheetRow}: no valid assignee emails resolved.`);
        setAutoNote(checkboxRange, 'No assignees could be matched in the Directory tab.');
        continue;
      }

      // --- Build calendar event resource ---
      const summary = project ? `[${project}] ${taskName}` : taskName;
      const startDate = formatDateISO(deadline);
      const endDateObj = new Date(deadline);
      endDateObj.setDate(endDateObj.getDate() + 1);
      const endDate = formatDateISO(endDateObj);

      const eventResource = {
        summary: summary,
        start: { date: startDate },
        end:   { date: endDate },
        attendees: emails.map(function(e) { return { email: e }; })
      };

      // --- Create or update ---
      if (existingId) {
        console.log(`Row ${sheetRow}: updating event for "${taskName}".`);
        try {
          Calendar.Events.patch(eventResource, CALENDAR_ID, existingId, { sendUpdates: 'all' });
        }
        catch (patchErr) {
          if (String(patchErr).indexOf('Not Found') !== -1) {
            console.log(`Row ${sheetRow}: event not found — creating fresh.`);
            const freshEvent = Calendar.Events.insert(eventResource, CALENDAR_ID, { sendUpdates: 'all' });
            tasksSheet.getRange(sheetRow, col.eventId + 1).setValue(freshEvent.id);
          }
          else {
            throw patchErr;
          }
        }
        updated++;
      }
      else {
        console.log(`Row ${sheetRow}: creating new event for "${taskName}".`);
        const newEvent = Calendar.Events.insert(eventResource, CALENDAR_ID, { sendUpdates: 'all' });
        tasksSheet.getRange(sheetRow, col.eventId + 1).setValue(newEvent.id);

        if (!String(row[col.status] || '').trim()) {
          tasksSheet.getRange(sheetRow, col.status + 1).setValue(TASK_STATUS.NOT_STARTED);
        }
        created++;
      }

      // --- Success — clear any auto-note and uncheck ---
      clearAutoNote(checkboxRange);
      checkboxRange.setValue(false);

    }
    catch (err) {
      console.error(`Row ${sheetRow}: ${err.message}`);
      setAutoNote(checkboxRange, 'Code error; please contact your admin.');
    }
  }

  console.log(`processTask: complete — ${created} created, ${updated} updated.`);
}

// --------------- Daily Reminders (~8:00 AM trigger) ---------------

/**
 * Sends reminder emails for tasks approaching their deadline.
 * One email per task, with all assignees as recipients.
 */
function sendReminders() {
  const env = validateSetup();
  const { tasksSheet, taskColMap } = env;
  const directory = loadDirectory(env.directorySheet, env.directoryColMap);
  const projectStatuses = loadProjectStatusMap(env.projectsSheet, env.projectColMap);
  const spreadsheetUrl = env.ss.getUrl();

  const data = tasksSheet.getDataRange().getValues();

  const col = {
    taskName: taskColMap[TASK_COLUMNS.TASK_NAME],
    project:  taskColMap[TASK_COLUMNS.PROJECT],
    assignee: taskColMap[TASK_COLUMNS.ASSIGNEE],
    deadline: taskColMap[TASK_COLUMNS.DEADLINE],
    status:   taskColMap[TASK_COLUMNS.STATUS],
    eventId:  taskColMap[TASK_COLUMNS.CALENDAR_EVENT_ID],
    notes:    taskColMap[TASK_COLUMNS.NOTES]
  };

  console.log('sendReminders: starting.');
  let sent = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const sheetRow = i + 1;

    try {
      if (String(row[col.status] || '').trim() === TASK_STATUS.COMPLETE) continue;

      const project = String(row[col.project] || '').trim();
      if (project) {
        const projStatus = projectStatuses.get(project);
        if (projStatus === PROJECT_STATUS.COMPLETE || projStatus === PROJECT_STATUS.CANCELLED) continue;
      }

      if (!String(row[col.eventId] || '').trim()) continue;

      const deadline = row[col.deadline];
      if (!(deadline instanceof Date)) continue;

      const days = daysUntil(deadline);
      if (!REMINDER_DAYS.includes(days)) continue;

      const taskName    = String(row[col.taskName] || '').trim();
      const assigneeStr = String(row[col.assignee] || '').trim();
      const emails = resolveAssigneeEmails(assigneeStr, directory);
      if (emails.length === 0) continue;

      const notes = String(row[col.notes] || '').trim();
      const notesRow = notes
        ? '<tr><td style="padding: 8px 12px; font-weight: bold; background: #f5f5f5;">Notes</td>'
          + '<td style="padding: 8px 12px; background: #f5f5f5; word-break: break-word;">'
          + escapeHtml(notes) + '</td></tr>'
        : '';

      const tokens = {
        TASK_NAME:      taskName,
        PROJECT_NAME:   project || '(none)',
        DEADLINE:       formatDateReadable(deadline),
        DAYS_UNTIL_DUE: String(days),
        NOTES_ROW:      notesRow,
        SPREADSHEET_URL: spreadsheetUrl
      };

      const subject = substituteTokens(EMAIL_TEMPLATES.REMINDER.SUBJECT, tokens);
      const body    = substituteTokens(EMAIL_TEMPLATES.REMINDER.BODY, Object.assign({}, tokens, {
        TASK_NAME:    escapeHtml(taskName),
        PROJECT_NAME: escapeHtml(project || '(none)')
      }));

      GmailApp.sendEmail(emails.join(','), subject, '', { htmlBody: body });
      console.log(`Row ${sheetRow}: reminder sent for "${taskName}" (due in ${days} days, ${emails.length} recipient(s)).`);
      sent++;

    }
    catch (err) {
      console.error(`Reminder error on row ${sheetRow}: ${err.message}`);
    }
  }

  console.log(`sendReminders: complete — ${sent} reminder(s) sent.`);
}
