/**
 * PermissionService class - Handles syncing sharing permissions based on
 * Directory roles and the new Global Access / Main File Role / Project Scope model.
 *
 * It synchronizes permissions for:
 * - The main spreadsheet (Project Directory)
 * - The root instance folder
 * - The Project Folders parent folder
 * - Each individual project folder
 */
class PermissionService {
  /**
   * Creates a new PermissionService instance.
   * @param {ExecutionContext} ctx - The execution context
   */
  constructor(ctx) {
    this.ctx = ctx;
    this.config = ctx.config;
    this.directory = ctx.directory;
    this.projectSheet = ctx.projectSheet;
    this.notificationService = ctx.notificationService;
  }

  /**
   * Refreshes sharing permissions for the spreadsheet and all related folders.
   * Called from the public refreshPermissions entrypoint.
   */
  refreshAllPermissions() {
    const permissionErrors = [];

    const spreadsheetId = this.ctx.spreadsheetId;
    let spreadsheetFile;

    try {
      spreadsheetFile = withBackoff(() => DriveApp.getFileById(spreadsheetId));
    } catch (e) {
      permissionErrors.push(`Cannot access main spreadsheet (ID: ${spreadsheetId}): ${e.message}`);
      this.notifyIfErrors(permissionErrors, spreadsheetId);
      return;
    }

    const owner = spreadsheetFile.getOwner();
    const spreadsheetOwnerEmail = owner ? owner.getEmail().toLowerCase() : '';

    // Precompute access rows from Directory (one per staff row)
    const accessRows = this.directory.getAccessRows();

    // 1. Main spreadsheet permissions
    try {
      const { desiredRoles, removeEmails } = this.buildSpreadsheetPermissions(accessRows);
      this.syncDrivePermissions(
        spreadsheetFile,
        desiredRoles,
        removeEmails,
        spreadsheetOwnerEmail,
        'main spreadsheet',
        permissionErrors
      );
    } catch (e) {
      permissionErrors.push(`Unexpected error while syncing main spreadsheet permissions: ${e.message}`);
    }

    // 2. Project Folders parent folder permissions (Project Scope baseline)
    const parentFolderId = this.config.parentFolderId;
    if (!parentFolderId) {
      permissionErrors.push('Parent Folder ID is not configured in Config sheet.');
    } else {
      try {
        const parentFolder = withBackoff(() => DriveApp.getFolderById(parentFolderId));
        const parentOwner = parentFolder.getOwner();
        const parentOwnerEmail = parentOwner ? parentOwner.getEmail().toLowerCase() : spreadsheetOwnerEmail;

        const { desiredRoles, removeEmails } = this.buildProjectsParentPermissions(accessRows);
        this.syncDrivePermissions(
          parentFolder,
          desiredRoles,
          removeEmails,
          parentOwnerEmail,
          'Project Folders parent',
          permissionErrors
        );
      } catch (e) {
        permissionErrors.push(`Cannot access Project Folders parent (ID: ${parentFolderId}): ${e.message}`);
      }
    }

    // 3. Root folder permissions (Global Access only)
    const rootFolderId = this.config.rootFolderId;
    if (!rootFolderId) {
      permissionErrors.push('Root Folder ID is not configured in Config sheet.');
    } else {
      try {
        const rootFolder = withBackoff(() => DriveApp.getFolderById(rootFolderId));
        const rootOwner = rootFolder.getOwner();
        const rootOwnerEmail = rootOwner ? rootOwner.getEmail().toLowerCase() : spreadsheetOwnerEmail;

        const { desiredRoles, removeEmails } = this.buildRootFolderPermissions(accessRows);
        this.syncDrivePermissions(
          rootFolder,
          desiredRoles,
          removeEmails,
          rootOwnerEmail,
          'root folder',
          permissionErrors
        );
      } catch (e) {
        permissionErrors.push(`Cannot access Root Folder (ID: ${rootFolderId}): ${e.message}`);
      }
    }

    // 4. Individual project folders - re-apply explicit sharing for assignees/requesters
    try {
      this.refreshProjectFolderPermissions(permissionErrors);
    } catch (e) {
      permissionErrors.push(`Unexpected error while refreshing individual project folder permissions: ${e.message}`);
    }

    this.notifyIfErrors(permissionErrors, spreadsheetId);
  }

  /**
   * Refreshes explicit sharing on each individual project folder.
   * Re-applies sharing for project assignees and requesters so they retain access
   * even if their Global/parent folder access was removed.
   *
   * This is designed to run AFTER parent/root folder permission changes so any
   * inherited access removals are applied first, and then explicit project-level
   * access is restored where appropriate.
   *
   * @param {string[]} permissionErrors - Array to append error messages into
   */
  refreshProjectFolderPermissions(permissionErrors) {
    const projects = this.projectSheet && typeof this.projectSheet.getProjects === 'function'
      ? this.projectSheet.getProjects()
      : [];

    if (!projects || projects.length === 0) {
      DEBUG && console.log('PermissionService: No projects found while refreshing project folder permissions');
      return;
    }

    DEBUG && console.log(`PermissionService: Refreshing explicit sharing for ${projects.length} project folder(s)`);

    for (const project of projects) {
      const folderId = project && typeof project.folderId === 'string'
        ? project.folderId
        : project && project.folderId;

      if (!folderId) {
        continue;
      }

      try {
        // Reuse ProjectService logic which already:
        // - Resolves assignee/requester names to emails
        // - Uses Advanced Drive API to add writers without notification emails
        // - Sends its own detailed error notifications on partial failures
        this.ctx.projectService.shareProjectFolder(project, folderId);
      } catch (e) {
        const projectId = project && typeof project.projectId === 'string'
          ? project.projectId
          : (project && project.projectId);

        permissionErrors.push(
          `Could not refresh sharing for project folder (Project ID: ${projectId || '(unknown)'}, Folder ID: ${folderId}): ${e.message}`
        );
      }
    }
  }

  /**
   * Builds desired roles and removals for the main spreadsheet.
   * @param {Object[]} accessRows - Directory access rows
   * @returns {{desiredRoles: Map<string, string>, removeEmails: Set<string>}}
   */
  buildSpreadsheetPermissions(accessRows) {
    const desiredRoles = new Map();
    const removeEmails = new Set();

    for (const row of accessRows) {
      const email = row.email;
      const activeFlag = row.activeFlag;
      const role = row.effectiveSpreadsheetRole;

      if (activeFlag === 'blank') {
        // Not managed yet - no changes
        continue;
      }

      if (activeFlag === 'no') {
        // Explicitly disabled - remove any existing access
        removeEmails.add(email);
        continue;
      }

      // Active = yes
      if (role === DIRECTORY_ACCESS_ROLES.EDITOR) {
        desiredRoles.set(email, 'edit');
      } else if (role === DIRECTORY_ACCESS_ROLES.VIEWER) {
        desiredRoles.set(email, 'view');
      } else {
        // Managed but should not have spreadsheet access
        removeEmails.add(email);
      }
    }

    return { desiredRoles, removeEmails };
  }

  /**
   * Builds desired roles and removals for the root folder.
   * Root sharing is driven solely by Global Access.
   * @param {Object[]} accessRows - Directory access rows
   * @returns {{desiredRoles: Map<string, string>, removeEmails: Set<string>}}
   */
  buildRootFolderPermissions(accessRows) {
    const desiredRoles = new Map();
    const removeEmails = new Set();

    for (const row of accessRows) {
      const email = row.email;
      const activeFlag = row.activeFlag;
      const globalRole = row.globalAccessRole;

      if (activeFlag === 'blank') {
        continue;
      }

      if (activeFlag === 'no') {
        removeEmails.add(email);
        continue;
      }

      // Active = yes
      if (globalRole === DIRECTORY_ACCESS_ROLES.EDITOR) {
        desiredRoles.set(email, 'edit');
      } else if (globalRole === DIRECTORY_ACCESS_ROLES.VIEWER) {
        desiredRoles.set(email, 'view');
      } else {
        // Managed staff with no Global Access should not have root access
        removeEmails.add(email);
      }
    }

    return { desiredRoles, removeEmails };
  }

  /**
   * Builds desired roles and removals for the Project Folders parent folder.
   * Uses effective folder scope as baseline (All-Editor / All-Viewer / Assigned-only).
   * @param {Object[]} accessRows - Directory access rows
   * @returns {{desiredRoles: Map<string, string>, removeEmails: Set<string>}}
   */
  buildProjectsParentPermissions(accessRows) {
    const desiredRoles = new Map();
    const removeEmails = new Set();

    for (const row of accessRows) {
      const email = row.email;
      const activeFlag = row.activeFlag;
      const scope = row.effectiveFolderScope;

      if (activeFlag === 'blank') {
        continue;
      }

      if (activeFlag === 'no') {
        removeEmails.add(email);
        continue;
      }

      if (scope === DIRECTORY_FOLDER_SCOPES.ALL_EDITOR) {
        desiredRoles.set(email, 'edit');
      } else if (scope === DIRECTORY_FOLDER_SCOPES.ALL_VIEWER) {
        desiredRoles.set(email, 'view');
      } else {
        // Assigned-only / none: no direct access to parent folder
        removeEmails.add(email);
      }
    }

    return { desiredRoles, removeEmails };
  }

  /**
   * Synchronizes Drive permissions on a file or folder to match the desired roles.
   * Only modifies explicitly managed emails; others are left unchanged.
   *
   * @param {GoogleAppsScript.Drive.File|GoogleAppsScript.Drive.Folder} entity - File or folder
   * @param {Map<string, string>} desiredRoles - Map of email -> 'edit' | 'view'
   * @param {Set<string>} removeEmails - Emails that should have no access
   * @param {string} ownerEmail - Owner email (never removed)
   * @param {string} label - Human-readable label for logging
   * @param {string[]} permissionErrors - Array to append error messages into
   */
  syncDrivePermissions(entity, desiredRoles, removeEmails, ownerEmail, label, permissionErrors) {
    const owner = (ownerEmail || '').toLowerCase();

    let currentEditors = [];
    let currentViewers = [];

    try {
      currentEditors = entity.getEditors().map(e => e.getEmail().toLowerCase());
      currentViewers = entity.getViewers().map(v => v.getEmail().toLowerCase());
    } catch (e) {
      permissionErrors.push(`Could not read existing permissions for ${label}: ${e.message}`);
      return;
    }

    // Additions and upgrades/downgrades
    for (const [emailRaw, desiredPerm] of desiredRoles) {
      const email = String(emailRaw || '').toLowerCase();
      if (!email || email === owner) {
        continue;
      }

      const isEditor = currentEditors.includes(email);
      const isViewer = currentViewers.includes(email);

      if (desiredPerm === 'edit') {
        if (!isEditor) {
          try {
            if (isViewer) {
              entity.removeViewer(email);
            }
            entity.addEditor(email);
            DEBUG && console.log(`PermissionService: Set editor on ${label}: ${email}`);
          } catch (e) {
            permissionErrors.push(`Could not set editor on ${label} for ${email}: ${e.message}`);
          }
        }
      } else if (desiredPerm === 'view') {
        if (isEditor) {
          try {
            entity.removeEditor(email);
            entity.addViewer(email);
            DEBUG && console.log(`PermissionService: Downgraded to viewer on ${label}: ${email}`);
          } catch (e) {
            permissionErrors.push(`Could not downgrade editor to viewer on ${label} for ${email}: ${e.message}`);
          }
        } else if (!isViewer) {
          try {
            entity.addViewer(email);
            DEBUG && console.log(`PermissionService: Set viewer on ${label}: ${email}`);
          } catch (e) {
            permissionErrors.push(`Could not add viewer on ${label} for ${email}: ${e.message}`);
          }
        }
      }
    }

    // Explicit removals
    for (const emailRaw of removeEmails) {
      const email = String(emailRaw || '').toLowerCase();
      if (!email || email === owner) {
        continue;
      }

      const isEditor = currentEditors.includes(email);
      const isViewer = currentViewers.includes(email);

      if (!isEditor && !isViewer) {
        continue;
      }

      try {
        if (isEditor) {
          entity.removeEditor(email);
        }
        if (isViewer) {
          entity.removeViewer(email);
        }
        DEBUG && console.log(`PermissionService: Removed access on ${label}: ${email}`);
      } catch (e) {
        permissionErrors.push(`Could not remove access on ${label} for ${email}: ${e.message}`);
      }
    }
  }

  /**
   * Sends a consolidated error notification if any permission errors occurred.
   * @param {string[]} permissionErrors - Array of error messages
   * @param {string} spreadsheetId - The main spreadsheet ID
   */
  notifyIfErrors(permissionErrors, spreadsheetId) {
    if (!permissionErrors || permissionErrors.length === 0) {
      return;
    }

    const lines = [
      'One or more permission updates failed while refreshing sharing permissions.',
      '',
      `Spreadsheet ID: ${spreadsheetId}`,
      '',
      'Failures:'
    ];

    for (const msg of permissionErrors) {
      lines.push(`  • ${msg}`);
    }

    try {
      this.notificationService.sendErrorNotification(
        'Permission Refresh Issues',
        lines.join('\n')
      );
    } catch (notifyError) {
      console.error(`PermissionService: Failed to send permission error notification: ${notifyError.message}`);
    }
  }
}

