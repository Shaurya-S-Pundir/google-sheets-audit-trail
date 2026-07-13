/**
 * @fileoverview Snapshot.js — Snapshot cache management.
 *
 * Each tracked sheet has TWO hidden companion sheets:
 *  - "__snapshot__<SheetName>"   — stores the last-known cell VALUES.
 *  - "__snapformula__<SheetName>" — stores the last-known cell FORMULAS.
 *
 * By maintaining both, we can distinguish:
 *  a) A formula change   (formula text differs)
 *  b) A value-only change (formula same, computed value differs — NOT logged)
 *  c) A plain value edit  (no formula, raw value differs)
 *
 * All snapshot sheets are hidden so they don't clutter the user's tab bar.
 */

// ============================================================
// Public API — read
// ============================================================

/**
 * Returns the cached VALUES snapshot for a given sheet name.
 * Returns null if no snapshot exists yet.
 *
 * @param {string} sheetName - The name of the source sheet.
 * @return {Array[]|null} 2D array of values, or null.
 */
function getSnapshotValues(sheetName) {
  return _readSnapshot(SNAPSHOT_SHEET_PREFIX + sheetName);
}

/**
 * Returns the cached FORMULA snapshot for a given sheet name.
 * Returns null if no snapshot exists yet.
 *
 * @param {string} sheetName - The name of the source sheet.
 * @return {Array[]|null} 2D array of formula strings, or null.
 */
function getSnapshotFormulas(sheetName) {
  return _readSnapshot(SNAPSHOT_FORMULA_PREFIX + sheetName);
}

// ============================================================
// Public API — write
// ============================================================

/**
 * Performs a full initialisation of snapshots for a given sheet.
 * Reads the entire sheet and writes both value and formula snapshots.
 * Safe to call even if snapshots already exist (will overwrite).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to snapshot.
 */
function initializeSnapshot(sheet) {
  const name = sheet.getName();
  if (_isIgnoredSheet(name)) return;

  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastCol = Math.max(sheet.getLastColumn(), 1);

  const values   = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const formulas = sheet.getRange(1, 1, lastRow, lastCol).getFormulas();

  _writeSnapshot(SNAPSHOT_SHEET_PREFIX   + name, values);
  _writeSnapshot(SNAPSHOT_FORMULA_PREFIX + name, formulas);
}

/**
 * Updates only a sub-range of the value snapshot for a sheet.
 * Used after an edit/paste to keep the snapshot in sync without
 * re-reading the entire sheet.
 *
 * @param {string}   sheetName  - The name of the source sheet.
 * @param {number}   startRow   - 1-indexed start row of the updated range.
 * @param {number}   startCol   - 1-indexed start column of the updated range.
 * @param {Array[]}  values     - 2D values to write into the snapshot.
 * @param {Array[]}  formulas   - 2D formulas to write into the snapshot.
 */
function updateSnapshotRange(sheetName, startRow, startCol, values, formulas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _ensureSnapshotSheetCapacity(
    SNAPSHOT_SHEET_PREFIX   + sheetName, startRow, startCol, values
  );
  _ensureSnapshotSheetCapacity(
    SNAPSHOT_FORMULA_PREFIX + sheetName, startRow, startCol, formulas
  );

  const valSnap = ss.getSheetByName(SNAPSHOT_SHEET_PREFIX   + sheetName);
  const fmlSnap = ss.getSheetByName(SNAPSHOT_FORMULA_PREFIX + sheetName);

  if (valSnap && values.length > 0 && values[0].length > 0) {
    valSnap.getRange(startRow, startCol, values.length, values[0].length)
           .setValues(values);
  }
  if (fmlSnap && formulas.length > 0 && formulas[0].length > 0) {
    fmlSnap.getRange(startRow, startCol, formulas.length, formulas[0].length)
           .setValues(formulas);
  }
}

/**
 * Deletes both snapshot sheets for a given sheet name.
 * Call when a sheet is deleted.
 *
 * @param {string} sheetName - The deleted sheet's name.
 */
function deleteSnapshot(sheetName) {
  _deleteSnapshotSheet(SNAPSHOT_SHEET_PREFIX   + sheetName);
  _deleteSnapshotSheet(SNAPSHOT_FORMULA_PREFIX + sheetName);
}

/**
 * Renames both snapshot sheets when a source sheet is renamed.
 *
 * @param {string} oldName - Previous sheet name.
 * @param {string} newName - New sheet name.
 */
function renameSnapshot(oldName, newName) {
  _renameSnapshotSheet(SNAPSHOT_SHEET_PREFIX   + oldName,
                       SNAPSHOT_SHEET_PREFIX   + newName);
  _renameSnapshotSheet(SNAPSHOT_FORMULA_PREFIX + oldName,
                       SNAPSHOT_FORMULA_PREFIX + newName);
}

// ============================================================
// Private helpers
// ============================================================

/**
 * Reads all values from a snapshot sheet by its internal name.
 *
 * @param {string} snapSheetName - Internal snapshot sheet name.
 * @return {Array[]|null}
 */
function _readSnapshot(snapSheetName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(snapSheetName);
  if (!sheet) return null;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return [];

  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

/**
 * Writes a 2D array to a snapshot sheet. Creates the sheet if missing.
 * The sheet is hidden immediately after creation.
 *
 * @param {string}  snapSheetName - Internal snapshot sheet name.
 * @param {Array[]} data          - 2D array to write.
 */
function _writeSnapshot(snapSheetName, data) {
  if (!data || data.length === 0) return;

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(snapSheetName);

  if (!sheet) {
    sheet = ss.insertSheet(snapSheetName);
    sheet.hideSheet();
  }

  // Clear the sheet first, then write fresh data.
  sheet.clearContents();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}

/**
 * Ensures the snapshot sheet has enough rows/cols for the given range.
 * Expands the sheet by appending blank rows/cols if needed.
 *
 * @param {string}  snapSheetName
 * @param {number}  startRow
 * @param {number}  startCol
 * @param {Array[]} data          - 2D data that will be written (used for size).
 */
function _ensureSnapshotSheetCapacity(snapSheetName, startRow, startCol, data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(snapSheetName);

  if (!sheet) {
    sheet = ss.insertSheet(snapSheetName);
    sheet.hideSheet();
  }

  const neededRows = startRow + (data.length       || 0) - 1;
  const neededCols = startCol + (data[0]?.length   || 0) - 1;

  if (sheet.getMaxRows()    < neededRows) sheet.insertRowsAfter(sheet.getMaxRows(), neededRows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < neededCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), neededCols - sheet.getMaxColumns());
}

/**
 * Deletes a snapshot sheet by its internal name (if it exists).
 *
 * @param {string} snapSheetName
 */
function _deleteSnapshotSheet(snapSheetName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(snapSheetName);
  if (sheet) ss.deleteSheet(sheet);
}

/**
 * Renames a snapshot sheet.
 *
 * @param {string} oldSnapName
 * @param {string} newSnapName
 */
function _renameSnapshotSheet(oldSnapName, newSnapName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(oldSnapName);
  if (sheet) sheet.setName(newSnapName);
}

/**
 * Returns true if a sheet name should be excluded from snapshotting.
 *
 * @param {string} name - Sheet name to test.
 * @return {boolean}
 */
function _isIgnoredSheet(name) {
  return IGNORED_SHEET_PREFIXES.some(prefix => name.startsWith(prefix));
}
