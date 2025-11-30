/**
 * i2i Teaming Tool - Setup Wizard
 *
 * This script creates a complete district instance of the i2i Teaming Tool.
 * Run runSetupWizard() to create all necessary files and folders.
 *
 * BEFORE RUNNING:
 * 1. Ensure this script has edit access (temporarily granted)
 * 2. Run as the "robo school" account that should own the new files
 * 3. After completion, revoke edit access to this script
 */

// ===== SOURCE TEMPLATE IDS =====
// These are the master templates in the shared folder
const SOURCE_IDS = {
  MAIN_PROJECT_SHEET: '1tjSJ47eZhlx-7c4ILnCf2wOhu0CYBNPe0cMwiH3Z5os',
  PROJECT_FILE_TEMPLATE: '1ppKq07JYB17fACYrghuHxNdAgi5dieqLLFEVp_OvMf4',
  EMAIL_STATUS_CHANGE: '1UwUCIhI8t64DW6YvCvf_-4Ltckmuns6YutEvEVaOSqM',
  EMAIL_REMINDER: '1fyv4FXuabcwMB-qgTUosi4cIJE6QPD5_rJ9h4mKORgU',
  EMAIL_PROJECT_UPDATE: '1hnnU63VAquJ4YLbFGmdzMR1eKqA2s4XB1bZv4aKy1zI',
  EMAIL_PROJECT_CANCELLATION: '1oVgN_mpXjgQ1PG9ulNW1D_UdKF8p7NypgtJ_uXvXCJA',
  EMAIL_NEW_PROJECT: '139_bnhZMwNIZHxqRgTw2D_rVuFWx2r-9ZSSnEYFmnms',
  FORM_TEMPLATE: '1Fp94TUfL-Mgp4_2wWE9k-RYK1bHpcXHRE4vgTigourU'
};

// ===== FILE AND FOLDER NAMES =====
// Customize these to change the names of created files/folders
const NAMES = {
  ROOT_FOLDER: 'i2i Teaming Tool',
  MAIN_SPREADSHEET: 'Project Directory',
  FORM: 'Project Submission Form',
  TEMPLATES_FOLDER: 'Templates',
  PROJECT_FOLDERS: 'Project Folders',
  PROJECT_FILE_TEMPLATE: 'Project File Template',
  EMAIL_NEW_PROJECT: 'Email Template - New Project',
  EMAIL_REMINDER: 'Email Template - Reminder',
  EMAIL_STATUS_CHANGE: 'Email Template - Status Change',
  EMAIL_PROJECT_UPDATE: 'Email Template - Project Update',
  EMAIL_PROJECT_CANCELLATION: 'Email Template - Project Cancellation'
};

// ===== SHEET NAMES (must match library constants) =====
const SHEET_NAMES = {
  CONFIG: 'Config',
  FORM_RESPONSES: 'Form Responses (Raw)'
};

// ===== CONFIG KEYS (must match library constants) =====
const CONFIG_KEYS = {
  DISTRICT_ID: 'District ID',
  SCHOOL_YEAR: 'School Year',
  NEXT_SERIAL: 'Next Serial',
  PARENT_FOLDER_ID: 'Parent Folder ID',
  PROJECT_TEMPLATE_ID: 'Project Template ID',
  FORM_ID: 'Form ID',
  ERROR_EMAIL_ADDRESSES: 'Error Email Addresses',
  EMAIL_NEW_PROJECT: 'Email Template - New Project',
  EMAIL_REMINDER: 'Email Template - Reminder',
  EMAIL_STATUS_CHANGE: 'Email Template - Status Change',
  EMAIL_PROJECT_UPDATE: 'Email Template - Project Update',
  EMAIL_PROJECT_CANCELLATION: 'Email Template - Project Cancellation',
  DEBUG_MODE: 'Debug Mode'
};

// ===== MAIN ENTRY POINT =====

/**
 * Main setup wizard function. Creates a complete district instance.
 * Run this function to set up a new district.
 */
function runSetupWizard() {
  console.log('=== i2i Teaming Tool Setup Wizard ===');
  console.log('Starting setup...');

  const runnerEmail = Session.getActiveUser().getEmail();
  console.log(`Running as: ${runnerEmail}`);

  // Check for existing setup folders (warn about duplicates)
  const existingFolders = checkExistingSetup();
  if (existingFolders.length > 0) {
    console.warn(`Warning: Found ${existingFolders.length} existing setup folder(s)`);
    for (const folder of existingFolders) {
      console.warn(`  - ${folder.getName()} (${folder.getId()})`);
    }
  }

  try {
    // Step 1: Create folder structure
    console.log('\n--- Step 1: Creating folder structure ---');
    const folders = createFolderStructure();

    // Step 2: Copy main spreadsheet (includes bound client script)
    console.log('\n--- Step 2: Copying main spreadsheet ---');
    const spreadsheet = copyMainSpreadsheet(folders.root);

    // Step 3: Copy form and link to spreadsheet
    console.log('\n--- Step 3: Setting up form ---');
    const formInfo = setupForm(folders.root, spreadsheet);

    // Step 4: Copy templates to Templates folder
    console.log('\n--- Step 4: Copying templates ---');
    const templateIds = copyTemplates(folders.templates);

    // Step 5: Update Config sheet
    console.log('\n--- Step 5: Updating Config sheet ---');
    updateConfigSheet(spreadsheet, {
      parentFolderId: folders.projectFolders.getId(),
      projectTemplateId: templateIds.projectFile,
      formId: formInfo.formId,
      errorEmail: runnerEmail,
      emailTemplates: templateIds.emailTemplates
    });

    // Step 6: Send summary email
    console.log('\n--- Step 6: Sending summary email ---');
    sendSetupSummaryEmail(runnerEmail, {
      rootFolder: folders.root,
      spreadsheet: spreadsheet,
      formUrl: formInfo.publishedUrl,
      formEditUrl: formInfo.editUrl
    });

    // TODO: Step 7: Apply sheet protections
    // applySheetProtections(spreadsheet);

    // Show completion message
    const message = `Setup complete!\n\n` +
      `Root Folder: ${folders.root.getName()}\n` +
      `Spreadsheet: ${spreadsheet.getName()}\n` +
      `Form URL: ${formInfo.publishedUrl}\n\n` +
      `A summary email has been sent to ${runnerEmail}.\n\n` +
      `NEXT STEPS:\n` +
      `1. Open the Config sheet and fill in District ID and School Year\n` +
      `2. Add staff to the Directory sheet\n` +
      `3. Run "Setup Triggers" from the Apps Script editor`;

    console.log('\n=== Setup Complete ===');
    console.log(message);

    // Try to show UI alert (only works if run from a UI context)
    try {
      SpreadsheetApp.getUi().alert('Setup Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      // Not in UI context, that's fine
    }

    return {
      success: true,
      rootFolderId: folders.root.getId(),
      spreadsheetId: spreadsheet.getId(),
      formId: formInfo.formId,
      formUrl: formInfo.publishedUrl
    };

  } catch (error) {
    console.error(`Setup failed: ${error.message}`);
    console.error(error.stack);

    // Try to show error in UI
    try {
      SpreadsheetApp.getUi().alert('Setup Failed', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      // Not in UI context
    }

    throw error;
  }
}

// ===== FOLDER STRUCTURE =====

/**
 * Creates the folder structure for a new district instance.
 * @returns {Object} Object containing folder references
 */
function createFolderStructure() {
  const dateStr = formatDate(new Date());
  const rootFolderName = `${NAMES.ROOT_FOLDER} - ${dateStr}`;

  console.log(`Creating root folder: ${rootFolderName}`);
  const rootFolder = DriveApp.createFolder(rootFolderName);

  console.log(`Creating subfolder: ${NAMES.TEMPLATES_FOLDER}`);
  const templatesFolder = rootFolder.createFolder(NAMES.TEMPLATES_FOLDER);

  console.log(`Creating subfolder: ${NAMES.PROJECT_FOLDERS}`);
  const projectFoldersFolder = rootFolder.createFolder(NAMES.PROJECT_FOLDERS);

  return {
    root: rootFolder,
    templates: templatesFolder,
    projectFolders: projectFoldersFolder
  };
}

/**
 * Checks for existing setup folders in the user's Drive root.
 * @returns {GoogleAppsScript.Drive.Folder[]} Array of matching folders
 */
function checkExistingSetup() {
  const folders = [];
  const iterator = DriveApp.getRootFolder().getFolders();

  while (iterator.hasNext()) {
    const folder = iterator.next();
    if (folder.getName().startsWith(NAMES.ROOT_FOLDER)) {
      folders.push(folder);
    }
  }

  return folders;
}

// ===== FILE COPYING =====

/**
 * Copies the main project spreadsheet to the target folder.
 * The copy includes the bound client script with library connection.
 * @param {GoogleAppsScript.Drive.Folder} targetFolder - Destination folder
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} The copied spreadsheet
 */
function copyMainSpreadsheet(targetFolder) {
  const sourceFile = DriveApp.getFileById(SOURCE_IDS.MAIN_PROJECT_SHEET);
  console.log(`Copying: ${sourceFile.getName()} → ${NAMES.MAIN_SPREADSHEET}`);

  const copiedFile = sourceFile.makeCopy(NAMES.MAIN_SPREADSHEET, targetFolder);
  const spreadsheet = SpreadsheetApp.openById(copiedFile.getId());

  // Hide the Status Snapshot sheet
  const snapshotSheet = spreadsheet.getSheetByName('Status Snapshot');
  if (snapshotSheet) {
    console.log('Hiding sheet: Status Snapshot');
    snapshotSheet.hideSheet();
  }

  console.log(`Created spreadsheet: ${spreadsheet.getId()}`);
  return spreadsheet;
}

/**
 * Copies the form template, links it to the spreadsheet, and publishes it.
 * @param {GoogleAppsScript.Drive.Folder} targetFolder - Destination folder
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - Target spreadsheet
 * @returns {Object} Form info including IDs and URLs
 */
function setupForm(targetFolder, spreadsheet) {
  // Copy the form
  const sourceFile = DriveApp.getFileById(SOURCE_IDS.FORM_TEMPLATE);
  console.log(`Copying: ${sourceFile.getName()} → ${NAMES.FORM}`);

  const copiedFile = sourceFile.makeCopy(NAMES.FORM, targetFolder);
  const form = FormApp.openById(copiedFile.getId());

  // Delete the placeholder "Form Responses (Raw)" sheet from the template copy
  // (the form will create its own responses sheet when linked)
  const existingResponsesSheet = spreadsheet.getSheetByName(SHEET_NAMES.FORM_RESPONSES);
  if (existingResponsesSheet) {
    console.log(`Deleting placeholder sheet: ${SHEET_NAMES.FORM_RESPONSES}`);
    spreadsheet.deleteSheet(existingResponsesSheet);
  }

  // Publish the form first (copied forms are in unpublished/draft state)
  // This must be done before setAcceptingResponses() or setDestination()
  console.log('Publishing form...');
  form.setPublished(true);

  // Link form to the spreadsheet (creates a new responses sheet)
  console.log('Linking form to spreadsheet...');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());

  // Small delay to let the linking complete
  Utilities.sleep(1000);

  // Enable response collection
  console.log('Enabling form responses...');
  form.setAcceptingResponses(true);

  // Give the API a moment to create the responses sheet
  SpreadsheetApp.flush();
  Utilities.sleep(2000);

  // Find and rename the form's responses sheet to match what the library expects
  const formResponsesSheet = findFormResponsesSheet(spreadsheet);
  if (formResponsesSheet && formResponsesSheet.getName() !== SHEET_NAMES.FORM_RESPONSES) {
    console.log(`Renaming "${formResponsesSheet.getName()}" → "${SHEET_NAMES.FORM_RESPONSES}"`);
    formResponsesSheet.setName(SHEET_NAMES.FORM_RESPONSES);
  }

  // Hide the Form Responses (Raw) sheet
  if (formResponsesSheet) {
    console.log(`Hiding sheet: ${SHEET_NAMES.FORM_RESPONSES}`);
    formResponsesSheet.hideSheet();
  }

  const formId = form.getId();
  const publishedUrl = form.getPublishedUrl();
  const editUrl = form.getEditUrl();

  console.log(`Form ID: ${formId}`);
  console.log(`Published URL: ${publishedUrl}`);

  return {
    formId: formId,
    publishedUrl: publishedUrl,
    editUrl: editUrl
  };
}

/**
 * Finds the form responses sheet in a spreadsheet.
 * Looks for sheets with names starting with "Form Responses".
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet to search
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} The form responses sheet, or null
 */
function findFormResponsesSheet(spreadsheet) {
  const sheets = spreadsheet.getSheets();
  for (const sheet of sheets) {
    const name = sheet.getName();
    // Match "Form Responses 1", "Form Responses 2", etc.
    if (name.startsWith('Form Responses') && name !== SHEET_NAMES.FORM_RESPONSES) {
      return sheet;
    }
  }
  return null;
}

/**
 * Copies all templates to the Templates folder.
 * @param {GoogleAppsScript.Drive.Folder} templatesFolder - Templates folder
 * @returns {Object} Object mapping template types to their new IDs
 */
function copyTemplates(templatesFolder) {
  const templateIds = {
    projectFile: null,
    emailTemplates: {}
  };

  // Copy project file template
  const projectTemplate = copyFileToFolder(
    SOURCE_IDS.PROJECT_FILE_TEMPLATE,
    NAMES.PROJECT_FILE_TEMPLATE,
    templatesFolder
  );
  templateIds.projectFile = projectTemplate.getId();

  // Copy email templates
  const emailTemplates = [
    { sourceId: SOURCE_IDS.EMAIL_NEW_PROJECT, name: NAMES.EMAIL_NEW_PROJECT, key: 'newProject' },
    { sourceId: SOURCE_IDS.EMAIL_REMINDER, name: NAMES.EMAIL_REMINDER, key: 'reminder' },
    { sourceId: SOURCE_IDS.EMAIL_STATUS_CHANGE, name: NAMES.EMAIL_STATUS_CHANGE, key: 'statusChange' },
    { sourceId: SOURCE_IDS.EMAIL_PROJECT_UPDATE, name: NAMES.EMAIL_PROJECT_UPDATE, key: 'projectUpdate' },
    { sourceId: SOURCE_IDS.EMAIL_PROJECT_CANCELLATION, name: NAMES.EMAIL_PROJECT_CANCELLATION, key: 'projectCancellation' }
  ];

  for (const template of emailTemplates) {
    const copiedFile = copyFileToFolder(template.sourceId, template.name, templatesFolder);
    templateIds.emailTemplates[template.key] = copiedFile.getId();
  }

  return templateIds;
}

/**
 * Copies a single file to a target folder.
 * @param {string} sourceId - Source file ID
 * @param {string} newName - Name for the copied file
 * @param {GoogleAppsScript.Drive.Folder} targetFolder - Destination folder
 * @returns {GoogleAppsScript.Drive.File} The copied file
 */
function copyFileToFolder(sourceId, newName, targetFolder) {
  const sourceFile = DriveApp.getFileById(sourceId);
  console.log(`Copying: ${sourceFile.getName()} → ${newName}`);

  const copiedFile = sourceFile.makeCopy(newName, targetFolder);
  console.log(`  Created: ${copiedFile.getId()}`);

  return copiedFile;
}

// ===== CONFIG SHEET UPDATE =====

/**
 * Updates the Config sheet with the new file/folder IDs.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet
 * @param {Object} config - Configuration values to set
 */
function updateConfigSheet(spreadsheet, config) {
  const configSheet = spreadsheet.getSheetByName(SHEET_NAMES.CONFIG);
  if (!configSheet) {
    throw new Error(`Config sheet "${SHEET_NAMES.CONFIG}" not found in spreadsheet`);
  }

  const data = configSheet.getDataRange().getValues();

  // Build key → row index map (Column A = key, Column B = value)
  const keyRowMap = new Map();
  for (let i = 0; i < data.length; i++) {
    const key = String(data[i][0] || '').trim();
    if (key) {
      keyRowMap.set(key, i + 1); // 1-based row number
    }
  }

  // Values to set
  const valuesToSet = [
    { key: CONFIG_KEYS.PARENT_FOLDER_ID, value: config.parentFolderId },
    { key: CONFIG_KEYS.PROJECT_TEMPLATE_ID, value: config.projectTemplateId },
    { key: CONFIG_KEYS.FORM_ID, value: config.formId },
    { key: CONFIG_KEYS.ERROR_EMAIL_ADDRESSES, value: config.errorEmail },
    { key: CONFIG_KEYS.EMAIL_NEW_PROJECT, value: config.emailTemplates.newProject },
    { key: CONFIG_KEYS.EMAIL_REMINDER, value: config.emailTemplates.reminder },
    { key: CONFIG_KEYS.EMAIL_STATUS_CHANGE, value: config.emailTemplates.statusChange },
    { key: CONFIG_KEYS.EMAIL_PROJECT_UPDATE, value: config.emailTemplates.projectUpdate },
    { key: CONFIG_KEYS.EMAIL_PROJECT_CANCELLATION, value: config.emailTemplates.projectCancellation },
    { key: CONFIG_KEYS.DEBUG_MODE, value: 'true' }
  ];

  for (const { key, value } of valuesToSet) {
    const row = keyRowMap.get(key);
    if (row) {
      configSheet.getRange(row, 2).setValue(value);
      console.log(`Set ${key} = ${value}`);
    } else {
      console.warn(`Config key "${key}" not found in Config sheet`);
    }
  }

  SpreadsheetApp.flush();
  console.log('Config sheet updated');
}

// ===== EMAIL =====

/**
 * Sends a setup summary email with next steps.
 * @param {string} recipientEmail - Email address to send to
 * @param {Object} info - Setup information
 */
function sendSetupSummaryEmail(recipientEmail, info) {
  const subject = `i2i Teaming Tool Setup Complete - ${info.rootFolder.getName()}`;
  const htmlBody = getSetupSummaryEmailHtml(info);

  GmailApp.sendEmail(recipientEmail, subject, '', {
    htmlBody: htmlBody,
    name: 'i2i Teaming Tool Setup Wizard'
  });

  console.log(`Summary email sent to ${recipientEmail}`);
}

/**
 * Generates the HTML body for the setup summary email.
 * @param {Object} info - Setup information
 * @returns {string} HTML email body
 */
function getSetupSummaryEmailHtml(info) {
  const folderUrl = info.rootFolder.getUrl();
  const spreadsheetUrl = info.spreadsheet.getUrl();

  return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; }
    h1 { color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 10px; }
    h2 { color: #202124; margin-top: 24px; }
    .success-banner { background: #e6f4ea; border-left: 4px solid #34a853; padding: 12px 16px; margin: 16px 0; }
    .info-box { background: #f8f9fa; border: 1px solid #dadce0; border-radius: 8px; padding: 16px; margin: 16px 0; }
    .info-box dt { font-weight: bold; color: #5f6368; margin-top: 8px; }
    .info-box dd { margin-left: 0; margin-bottom: 8px; }
    .url-box { background: #e8f0fe; padding: 12px; border-radius: 4px; word-break: break-all; margin: 8px 0; }
    .checklist { list-style: none; padding-left: 0; }
    .checklist li { padding: 8px 0; padding-left: 28px; position: relative; }
    .checklist li::before { content: "☐"; position: absolute; left: 0; color: #1a73e8; font-size: 18px; }
    .warning { background: #fef7e0; border-left: 4px solid #f9ab00; padding: 12px 16px; margin: 16px 0; }
    a { color: #1a73e8; }
    code { background: #f1f3f4; padding: 2px 6px; border-radius: 4px; font-family: monospace; }
  </style>
</head>
<body>
  <h1>i2i Teaming Tool Setup Complete</h1>
  
  <div class="success-banner">
    <strong>Your new district instance has been created successfully!</strong>
  </div>

  <h2>Created Resources</h2>
  <div class="info-box">
    <dl>
      <dt>Root Folder</dt>
      <dd><a href="${folderUrl}">${info.rootFolder.getName()}</a></dd>
      
      <dt>Project Directory (Main Spreadsheet)</dt>
      <dd><a href="${spreadsheetUrl}">${info.spreadsheet.getName()}</a></dd>
      
      <dt>Form URL (for users to submit projects)</dt>
      <dd class="url-box"><a href="${info.formUrl}">${info.formUrl}</a></dd>
      
      <dt>Form Edit URL (for admins)</dt>
      <dd><a href="${info.formEditUrl}">Edit Form</a></dd>
    </dl>
  </div>

  <h2>✅ Next Steps Checklist</h2>
  <ul class="checklist">
    <li><strong>Configure District Settings</strong><br>
      Open the <a href="${spreadsheetUrl}">Project Directory</a> → <code>Config</code> sheet<br>
      Fill in: <code>District ID</code> (e.g., "NUSD") and <code>School Year</code> (e.g., "25_26")<br>
      Review <code>Error Email Addresses</code> (comma-separated list for admin notifications)</li>
    
    <li><strong>Add Staff Directory</strong><br>
      Go to the <code>Directory</code> sheet<br>
      Add staff members with their Name, Email Address, and Permissions<br>
      Then update permissions: <code>Teaming Tool</code> → <code>Refresh Permissions</code><br>
      This syncs spreadsheet access with Directory permissions</li>
    
    <li><strong>Sync Form Dropdowns</strong><br>
      From the spreadsheet menu: <code>Teaming Tool</code> → <code>Sync Form Dropdowns</code><br>
      This populates the form's assignee dropdown from your Directory</li>
    
    <li><strong>Set Up Automation Triggers</strong><br>
      Open the spreadsheet → <code>Extensions</code> → <code>Apps Script</code><br>
      Run the <code>setupTriggers</code> function once<br>
      This creates the automated processing schedules</li>
    
    <li><strong>Distribute the Form Link</strong><br>
      Share the form URL with users who need to submit projects</li>
  </ul>

  <div class="warning">
    <strong>⚠️ Important:</strong> Don't forget to run <code>setupTriggers</code> from the Apps Script editor! 
    Without this, automated processing won't work.
  </div>

  <h2>Sheet Overview</h2>
  <div class="info-box">
    <dl>
      <dt>Project Management Sheet</dt>
      <dd>Main data sheet where all projects are tracked</dd>
      
      <dt>Config</dt>
      <dd>District-specific settings and file IDs</dd>
      
      <dt>Directory</dt>
      <dd>Staff list for assignee dropdowns and notifications</dd>
      
      <dt>Codes</dt>
      <dd>Dropdown options for categories, statuses, and reminder timelines</dd>
      
      <dt>Status Snapshot</dt>
      <dd>System sheet for tracking status changes (do not edit)</dd>
      
      <dt>Form Responses (Raw)</dt>
      <dd>Raw form submissions (do not edit)</dd>
    </dl>
  </div>

  <p style="color: #5f6368; font-size: 12px; margin-top: 32px; border-top: 1px solid #dadce0; padding-top: 16px;">
    This email was automatically generated by the i2i Teaming Tool Setup Wizard.<br>
    Setup completed: ${new Date().toLocaleString()}
  </p>
</body>
</html>
`;
}

// ===== UTILITY FUNCTIONS =====

/**
 * Formats a date as "MMM DD, YYYY" (e.g., "Jan 15, 2025").
 * @param {Date} date - The date to format
 * @returns {string} Formatted date string
 */
function formatDate(date) {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const month = months[date.getMonth()];
  const day = date.getDate();
  const year = date.getFullYear();
  return `${month} ${day}, ${year}`;
}

// ===== SHEET PROTECTIONS (TODO) =====

/**
 * Applies sheet protections to the spreadsheet.
 * TODO: Implement this function to set up protections as specified in scratch/setup.md
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet
 */
function applySheetProtections(spreadsheet) {
  // TODO: Implement sheet protections
  // - Protect Row 2 (internal keys) in Projects sheet
  // - Protect system columns (automation_status, calendar_event_id, folder_id, etc.)
  // - Protect Config Column A (keys)
  // - Protect Directory/Codes header rows
  // - Protect entire Status Snapshot sheet
  console.log('TODO: Sheet protections not yet implemented');
}

// ===== MENU (for running from spreadsheet context) =====

/**
 * Creates a menu in the spreadsheet UI.
 * Only works if this script is bound to a spreadsheet.
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Setup Wizard')
      .addItem('Run Setup Wizard', 'runSetupWizard')
      .addToUi();
  } catch (e) {
    // Not bound to a spreadsheet, that's fine
  }
}
