/**
 * @fileoverview Setup.js — One-time initialization and trigger installation.
 *
 * Run `initializeAuditSystem()` once from the Apps Script IDE to:
 *   1. Create the Audit Log sheet (if missing).
 *   2. Delete any stale audit-system triggers and install fresh ones.
 *   3. Initialize snapshots for all existing sheets.
 *   4. Persist the current sheet name list for rename detection.
 *
 * This function is safe to re-run; it is fully idempotent.
 *
 * Additionally, `resetAuditSystem()` is provided to wipe all audit data,
 * snapshots, and triggers — useful for testing or a clean reinstall.
 */

// ============================================================
// Public API
// ============================================================

/**
 * Main entry point. Run this once from the Apps Script editor:
 *   Extensions → Apps Script → Run → initializeAuditSystem
 *
 * Required OAuth scopes (granted automatically on first run):
 *   - https://www.googleapis.com/auth/spreadsheets
 *   - https://www.googleapis.com/auth/script.scriptapp  (trigger management)
 *   - https://www.googleapis.com/auth/userinfo.email
 */
function initializeAuditSystem() {
  console.log('[AuditTrail] Starting initialization…');

  // Step 1: Ensure the Audit Log sheet exists.
  ensureAuditLogSheet();
  console.log('[AuditTrail] Audit Log sheet ready.');

  // Step 2: Install all required installable triggers.
  _installTriggers();
  console.log('[AuditTrail] Triggers installed.');

  // Step 3: Snapshot all existing sheets.
  _initializeAllSnapshots();
  console.log('[AuditTrail] Snapshots initialized.');

  // Step 4: Persist sheet names for rename detection.
  _persistSheetNames();
  console.log('[AuditTrail] Sheet names persisted.');

  // Step 5: Mark as initialized.
  PropertiesService.getScriptProperties()
                   .setProperty(PROP_INITIALIZED, 'true');

  SpreadsheetApp.getUi().alert(
    '✅ Audit Trail Active',
    'The audit logging system has been successfully initialized.\n\n' +
    'All changes to this spreadsheet will now be recorded in the ' +
    '"Audit Log" sheet.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  console.log('[AuditTrail] Initialization complete.');
}

/**
 * Removes all audit-system triggers, snapshots, and (optionally) the Audit Log
 * sheet. Use this for a clean reinstall or during testing.
 *
 * ⚠️ This will permanently delete all recorded audit log entries.
 */
function resetAuditSystem() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Reset Audit System',
    'This will DELETE all triggers, snapshots, and audit log entries.\n\n' +
    'Are you sure you want to continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    console.log('[AuditTrail] Reset cancelled by user.');
    return;
  }

  // Remove all triggers.
  _removeAllAuditTriggers();
  console.log('[AuditTrail] Triggers removed.');

  // Delete all snapshot sheets.
  _deleteAllSnapshotSheets();
  console.log('[AuditTrail] Snapshots deleted.');

  // Delete the Audit Log sheet.
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(AUDIT_LOG_SHEET_NAME);
  if (logSheet) ss.deleteSheet(logSheet);
  console.log('[AuditTrail] Audit Log sheet deleted.');

  // Clear Script Properties.
  PropertiesService.getScriptProperties().deleteAllProperties();
  console.log('[AuditTrail] Script properties cleared.');

  ui.alert('✅ Reset Complete', 'The audit system has been fully reset.', ui.ButtonSet.OK);
}

/**
 * Adds a custom menu to the spreadsheet UI so users can run setup from
 * the menu bar without opening the Apps Script editor.
 *
 * Wire this up as a simple onOpen trigger (not installable):
 *   function onOpen() { onMenuOpen(); }
 */
function onMenuOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔍 Audit Trail')
    .addItem('Initialize / Re-initialize', 'initializeAuditSystem')
    .addSeparator()
    .addItem('Reset (Delete all data)', 'resetAuditSystem')
    .addToUi();
}

// ============================================================
// Private — trigger management
// ============================================================

/**
 * Installs all required installable triggers for the active spreadsheet.
 * Removes any existing audit-system triggers first to avoid duplicates.
 */
function _installTriggers() {
  _removeAllAuditTriggers();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // onEdit — handles cell edits and pastes.
  ScriptApp.newTrigger('onEditTrigger')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // onChange — handles structural changes (rows, columns, sheets).
  ScriptApp.newTrigger('onChangeTrigger')
    .forSpreadsheet(ss)
    .onChange()
    .create();

  // onOpen — re-initialises missing snapshots on spreadsheet open.
  ScriptApp.newTrigger('onOpenTrigger')
    .forSpreadsheet(ss)
    .onOpen()
    .create();
}

/**
 * Deletes all triggers whose handler function belongs to this audit system.
 */
function _removeAllAuditTriggers() {
  const auditHandlers = new Set([
    'onEditTrigger',
    'onChangeTrigger',
    'onOpenTrigger',
  ]);

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (auditHandlers.has(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// ============================================================
// Private — snapshot initialisation
// ============================================================

/**
 * Initialises value and formula snapshots for every tracked sheet in the
 * active spreadsheet. Skips ignored sheets (audit log, existing snapshots).
 */
function _initializeAllSnapshots() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (!_isIgnoredSheet(name)) {
      console.log(`[AuditTrail] Snapshotting: "${name}"`);
      initializeSnapshot(sheet);
    }
  });
}

/**
 * Deletes all hidden snapshot sheets created by the audit system.
 */
function _deleteAllSnapshotSheets() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith(SNAPSHOT_SHEET_PREFIX) ||
        name.startsWith(SNAPSHOT_FORMULA_PREFIX)) {
      ss.deleteSheet(sheet);
    }
  });
}

/**
 * Persists the current list of non-ignored sheet names to ScriptProperties.
 * Defined here as a local copy so Setup.js is self-contained on first load
 * (Triggers.js may not have been registered yet).
 */
function _persistSheetNames() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const names = ss.getSheets()
    .map(s => s.getName())
    .filter(n => !_isIgnoredSheet(n));

  PropertiesService.getScriptProperties()
                   .setProperty(PROP_SHEET_NAMES, JSON.stringify(names));
}

/**
 * Returns true if a sheet name should be excluded from auditing.
 * Duplicated here so Setup.js works as a standalone entry point.
 *
 * @param {string} name
 * @return {boolean}
 */
function _isIgnoredSheet(name) {
  return IGNORED_SHEET_PREFIXES.some(prefix => name.startsWith(prefix));
}
