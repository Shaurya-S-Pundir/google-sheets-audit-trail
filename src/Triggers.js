/**
 * @fileoverview Triggers.js — All event handler functions.
 *
 * This file contains the top-level functions that Apps Script calls when
 * spreadsheet events fire. Each handler is responsible for:
 *   1. Validating the event (ignoring internal/audit sheets).
 *   2. Collecting change data (using Diff.js and Snapshot.js).
 *   3. Writing log entries (using Logger.js).
 *   4. Updating the snapshot so the next event has correct old values.
 *
 * Entry points (must be registered as installable triggers):
 *   - onEditTrigger(e)    → Spreadsheet > On edit
 *   - onChangeTrigger(e)  → Spreadsheet > On change
 *   - onOpenTrigger(e)    → Spreadsheet > On open
 *
 * A LockService is used around state-mutating operations to prevent race
 * conditions when multiple users edit simultaneously.
 */

// ============================================================
// Entry point: onEdit
// ============================================================

/**
 * Handles cell edits and pastes. Fires for every value change, including
 * multi-cell pastes.
 *
 * Apps Script provides e.oldValue and e.value reliably for single-cell edits.
 * For multi-cell pastes, these fields are absent — in that case we fall back
 * to a full snapshot diff of the affected range.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEditTrigger(e) {
  try {
    if (!e || !e.range) return;

    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();

    // Skip internal sheets (audit log, snapshots).
    if (_isIgnoredSheet(sheetName)) return;

    const user = _getActiveUser();
    const numCells = range.getNumRows() * range.getNumColumns();

    let changes = [];

    if (numCells === 1) {
      // ── Single-cell edit ──
      // Use the event values for speed; snapshot for formula comparison.
      changes = detectSingleCellChange(e);
    } else {
      // ── Multi-cell edit / Paste ──
      // Determine change type: PASTE if this looks like a paste operation,
      // otherwise EDIT (e.g. fill-down, find-replace).
      const changeType = _isPaste(e) ? CHANGE_TYPE.PASTE : CHANGE_TYPE.EDIT;
      changes = detectCellChanges(sheet, range, changeType);
    }

    if (changes.length === 0) {
      // Still update snapshot even if nothing changed (e.g. formatting only).
      _refreshSnapshotForRange(sheet, range);
      return;
    }

    // Build log rows.
    const rows = changes.map(ch => buildLogRow({
      user: user,
      sheetName: sheetName,
      cell: ch.cell,
      row: ch.row,
      col: ch.col,
      oldValue: ch.oldValue,
      newValue: ch.newValue,
      changeType: ch.changeType,
    }));

    // Write all entries in one batch.
    logEntries(rows);

    // Update the snapshot to reflect the new state.
    _refreshSnapshotForRange(sheet, range);

  } catch (err) {
    console.error('[AuditTrail] onEditTrigger error:', err.message, err.stack);
  }
}

// ============================================================
// Entry point: onChange
// ============================================================

/**
 * Handles structural changes: row/column insertions and deletions, and
 * sheet-level operations (create, delete, rename).
 *
 * Apps Script's onChange event provides an `e.changeType` string:
 *   INSERT_ROW, DELETE_ROW, INSERT_COLUMN, DELETE_COLUMN,
 *   INSERT_GRID (new sheet), REMOVE_GRID (deleted sheet),
 *   FORMAT, OTHER.
 *
 * Sheet renames surface as changeType === 'OTHER'. We detect them by
 * comparing the current sheet list against a stored list in ScriptProperties.
 *
 * @param {GoogleAppsScript.Events.SheetsOnChange} e
 */
function onChangeTrigger(e) {
  try {
    if (!e) return;

    const user = _getActiveUser();
    const changeType = e.changeType || 'OTHER';

    switch (changeType) {
      case 'INSERT_ROW':
      case 'DELETE_ROW':
      case 'INSERT_COLUMN':
      case 'DELETE_COLUMN':
        _handleStructuralChange(e, user, changeType);
        break;

      case 'INSERT_GRID':
        _handleSheetCreate(user);
        break;

      case 'REMOVE_GRID':
        _handleSheetDelete(user);
        break;

      case 'FORMAT':
        // Formatting-only changes — explicitly ignored per requirements.
        break;

      case 'OTHER':
        // Could be a sheet rename or another unclassified event.
        _handlePossibleRename(user);
        break;

      default:
        // Unknown change type — log as OTHER for observability.
        _logStructuralEntry(user, '', '', '', CHANGE_TYPE.OTHER);
        break;
    }

    // Refresh stored sheet name list after any structural change.
    _persistSheetNames();

  } catch (err) {
    console.error('[AuditTrail] onChangeTrigger error:', err.message, err.stack);
  }
}

// ============================================================
// Entry point: onOpen
// ============================================================

/**
 * Fires when any user opens the spreadsheet.
 * Re-initialises any missing snapshots (e.g. after the script was first
 * installed, or after a snapshot sheet was accidentally deleted).
 *
 * @param {GoogleAppsScript.Events.SheetsOnOpen} e
 */
function onOpenTrigger(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();

    sheets.forEach(sheet => {
      const name = sheet.getName();
      if (_isIgnoredSheet(name)) return;

      // Only initialise if a snapshot is genuinely missing.
      if (!getSnapshotValues(name)) {
        initializeSnapshot(sheet);
      }
    });

    _persistSheetNames();

  } catch (err) {
    console.error('[AuditTrail] onOpenTrigger error:', err.message, err.stack);
  }
}

// ============================================================
// Private — structural change handlers
// ============================================================

/**
 * Logs a row/column insertion or deletion.
 * After a structural change, all snapshots for the active sheet are
 * re-initialised because row/column offsets shift.
 *
 * @param {Object} e          - The onChange event.
 * @param {string} user       - Active user email.
 * @param {string} changeType - One of the INSERT_* / DELETE_* constants.
 */
function _handleStructuralChange(e, user, changeType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  if (!activeSheet || _isIgnoredSheet(activeSheet.getName())) return;

  _logStructuralEntry(user, activeSheet.getName(), '', '', changeType);

  // Re-snapshot the entire sheet because offsets have shifted.
  initializeSnapshot(activeSheet);
}

/**
 * Detects a newly created sheet and logs it.
 * Initialises a fresh snapshot for the new sheet.
 *
 * @param {string} user - Active user email.
 */
function _handleSheetCreate(user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentNames = _getCurrentSheetNames();
  const storedNames = _getStoredSheetNames();

  // The new sheet is in currentNames but not in storedNames.
  const newSheets = currentNames.filter(n => !storedNames.includes(n));

  newSheets.forEach(name => {
    if (_isIgnoredSheet(name)) return;

    _logStructuralEntry(user, name, '', '', CHANGE_TYPE.SHEET_CREATE);

    const sheet = ss.getSheetByName(name);
    if (sheet) initializeSnapshot(sheet);
  });
}

/**
 * Detects a deleted sheet and logs it.
 * Cleans up the orphaned snapshot sheets.
 *
 * @param {string} user - Active user email.
 */
function _handleSheetDelete(user) {
  const currentNames = _getCurrentSheetNames();
  const storedNames = _getStoredSheetNames();

  // Deleted sheets are in storedNames but not in currentNames.
  const deletedSheets = storedNames.filter(n => !currentNames.includes(n));

  deletedSheets.forEach(name => {
    if (_isIgnoredSheet(name)) return;

    _logStructuralEntry(user, name, '', '', CHANGE_TYPE.SHEET_DELETE);
    deleteSnapshot(name);
  });
}

/**
 * Detects sheet renames by comparing the current sheet list against the stored
 * list. Apps Script does not directly report old/new names in the event object.
 *
 * Heuristic: if the count is the same but a name changed, it's a rename.
 * If counts differ, it's a create/delete handled elsewhere.
 *
 * @param {string} user - Active user email.
 */
function _handlePossibleRename(user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentNames = _getCurrentSheetNames();
  const storedNames = _getStoredSheetNames();

  if (currentNames.length !== storedNames.length) return; // Not a rename.

  // Find names that disappeared and appeared.
  const removed = storedNames.filter(n => !currentNames.includes(n));
  const added = currentNames.filter(n => !storedNames.includes(n));

  if (removed.length !== 1 || added.length !== 1) return; // Ambiguous — skip.

  const oldName = removed[0];
  const newName = added[0];

  if (_isIgnoredSheet(oldName) || _isIgnoredSheet(newName)) return;

  _logStructuralEntry(user, newName, oldName, newName, CHANGE_TYPE.SHEET_RENAME);
  renameSnapshot(oldName, newName);
}

// ============================================================
// Private — snapshot helpers
// ============================================================

/**
 * Refreshes the snapshot for the exact range that was just edited.
 * This is much cheaper than re-snapshotting the entire sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 */
function _refreshSnapshotForRange(sheet, range) {
  const values = range.getValues();
  const formulas = range.getFormulas();
  updateSnapshotRange(
    sheet.getName(),
    range.getRow(),
    range.getColumn(),
    values,
    formulas
  );
}

// ============================================================
// Private — logging helpers
// ============================================================

/**
 * Logs a single structural-change entry (no cell-level detail).
 *
 * @param {string} user
 * @param {string} sheetName
 * @param {string} oldValue
 * @param {string} newValue
 * @param {string} changeType
 */
function _logStructuralEntry(user, sheetName, oldValue, newValue, changeType) {
  const row = buildLogRow({
    user: user,
    sheetName: sheetName,
    cell: '',
    row: '',
    col: '',
    oldValue: oldValue,
    newValue: newValue,
    changeType: changeType,
  });
  logEntries([row]);
}

// ============================================================
// Private — sheet name tracking (for rename detection)
// ============================================================

/**
 * Returns the list of all current (non-ignored) sheet names.
 * @return {string[]}
 */
function _getCurrentSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(s => s.getName())
    .filter(n => !_isIgnoredSheet(n));
}

/**
 * Returns the previously persisted list of sheet names from ScriptProperties.
 * Returns empty array if never persisted.
 * @return {string[]}
 */
function _getStoredSheetNames() {
  try {
    const stored = PropertiesService.getScriptProperties()
      .getProperty(PROP_SHEET_NAMES);
    return stored ? JSON.parse(stored) : [];
  } catch (_) {
    return [];
  }
}

/**
 * Persists the current list of non-ignored sheet names to ScriptProperties.
 */
function _persistSheetNames() {
  const names = _getCurrentSheetNames();
  PropertiesService.getScriptProperties()
    .setProperty(PROP_SHEET_NAMES, JSON.stringify(names));
}

// ============================================================
// Private — utilities
// ============================================================

/**
 * Returns true if the given sheet name should be excluded from auditing.
 * @param {string} name
 * @return {boolean}
 */
function _isIgnoredSheet(name) {
  return IGNORED_SHEET_PREFIXES.some(prefix => name.startsWith(prefix));
}

/**
 * Attempts to get the active user's email. Returns an empty string in
 * contexts where this is not permitted (e.g. simple triggers).
 * @return {string}
 */
function _getActiveUser() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (_) {
    return '';
  }
}

/**
 * Heuristically determines whether an edit event represents a paste.
 * A paste typically affects multiple cells, and e.oldValue is undefined.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 * @return {boolean}
 */
function _isPaste(e) {
  const numCells = e.range.getNumRows() * e.range.getNumColumns();
  // Multi-cell with no oldValue supplied → likely a paste.
  return numCells > 1 && e.oldValue === undefined;
}
