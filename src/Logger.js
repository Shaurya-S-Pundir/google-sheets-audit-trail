/**
 * @fileoverview Logger.js — Core audit log writer.
 *
 * Responsible for:
 *  - Ensuring the Audit Log sheet exists with correct headers.
 *  - Constructing individual log row objects.
 *  - Batch-writing multiple log rows in a single API call for efficiency.
 */

// ============================================================
// Public API
// ============================================================

/**
 * Ensures the Audit Log sheet exists. Creates it with styled, frozen headers
 * if it does not. Safe to call on every trigger invocation (idempotent).
 *
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The Audit Log sheet.
 */
function ensureAuditLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(AUDIT_LOG_SHEET_NAME);

  if (!logSheet) {
    logSheet = ss.insertSheet(AUDIT_LOG_SHEET_NAME);
    _applyHeaderRow(logSheet);
    // Move audit log to the last position so it's out of the way.
    ss.setActiveSheet(logSheet);
    ss.moveActiveSheet(ss.getNumSheets());
  }

  return logSheet;
}

/**
 * Builds a single log row array in the column order defined by LOG_HEADERS.
 *
 * @param {Object} params
 * @param {string} params.user         - Email of the acting user (may be empty).
 * @param {number} params.row          - 1-indexed row number of the changed cell.
 * @param {string} params.columnHeader - Header text from row 1 of the changed column.
 * @param {*}      params.oldValue     - Previous cell value.
 * @param {*}      params.newValue     - New cell value.
 * @return {Array} A row array ready to be appended.
 */
function buildLogRow(params) {
  return [
    new Date(),                          // Timestamp
    params.user || '',                   // User
    params.row || '',                    // Row
    params.columnHeader || '',           // Column (header from row 1)
    _formatValue(params.oldValue),       // Old Value
    _formatValue(params.newValue),       // New Value
  ];
}

/**
 * Appends an array of log rows to the Audit Log sheet in one batch write.
 * This is far more efficient than calling appendRow() individually.
 *
 * @param {Array[]} rows - Array of row arrays produced by buildLogRow().
 */
function logEntries(rows) {
  if (!rows || rows.length === 0) return;

  try {
    const logSheet = ensureAuditLogSheet();
    const lastRow = logSheet.getLastRow();

    // Write all rows at once using setValues for maximum performance.
    const startRow = lastRow + 1;
    logSheet
      .getRange(startRow, 1, rows.length, LOG_COLUMN_COUNT)
      .setValues(rows);

  } catch (err) {
    // Log the error to Apps Script's execution log — never throw from here,
    // as a logging failure must not surface as an error to the user.
    console.error('[AuditTrail] logEntries failed:', err.message);
  }
}

// ============================================================
// Private helpers
// ============================================================

/**
 * Applies styled, frozen header row to a newly created Audit Log sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function _applyHeaderRow(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, LOG_COLUMN_COUNT);
  headerRange.setValues([LOG_HEADERS]);

  // Style: dark background, white bold text, freeze the row.
  headerRange
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10);

  sheet.setFrozenRows(1);

  // Column widths for the simplified 6-column layout.
  sheet.setColumnWidth(1, 180); // Timestamp
  sheet.setColumnWidth(2, 200); // User
  sheet.setColumnWidth(3, 60);  // Row
  sheet.setColumnWidth(4, 180); // Column (header name)
  sheet.setColumnWidth(5, 160); // Old Value
  sheet.setColumnWidth(6, 160); // New Value
}

/**
 * Normalises a cell value for display in the audit log.
 * Converts null/undefined to empty string; booleans to their string form.
 *
 * @param {*} value - Raw cell value.
 * @return {string|number|boolean} Normalised value.
 */
function _formatValue(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object' && value instanceof Date) {
    return value.toISOString();
  }
  return value;
}
