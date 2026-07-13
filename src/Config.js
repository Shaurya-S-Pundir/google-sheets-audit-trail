/**
 * @fileoverview Config.js — Central configuration constants for the Audit Trail system.
 *
 * All shared constants, sheet names, column definitions, and tuneable limits
 * are defined here. Modify this file to customise behaviour without touching
 * trigger or business-logic code.
 */

// ============================================================
// Sheet identifiers
// ============================================================

/** The visible audit-log sheet name that users will see. */
const AUDIT_LOG_SHEET_NAME = 'Audit Log';

/**
 * Prefix used for hidden snapshot sheets.
 * e.g. "__snapshot__Sales" stores the previous state of the "Sales" sheet.
 */
const SNAPSHOT_SHEET_PREFIX = '__snapshot__';

/**
 * Prefix used for snapshot sheets that store formula text.
 * e.g. "__snapformula__Sales"
 */
const SNAPSHOT_FORMULA_PREFIX = '__snapformula__';

// ============================================================
// Sheets that must never be audited (to prevent recursion)
// ============================================================

/**
 * Any sheet whose name starts with one of these strings is ignored.
 * This prevents the system from auditing its own internal sheets.
 */
const IGNORED_SHEET_PREFIXES = [
  SNAPSHOT_SHEET_PREFIX,
  SNAPSHOT_FORMULA_PREFIX,
  AUDIT_LOG_SHEET_NAME,
];

// ============================================================
// Audit Log column layout (1-indexed)
// ============================================================

/** Ordered list of column headers in the Audit Log sheet. */
const LOG_HEADERS = [
  'Timestamp',
  'User',
  'Row',
  'Column',
  'Old Value',
  'New Value',
];

/** Total number of log columns — derived from LOG_HEADERS for safety. */
const LOG_COLUMN_COUNT = LOG_HEADERS.length;

// ============================================================
// Change type constants
// ============================================================

const CHANGE_TYPE = {
  EDIT:             'EDIT',
  PASTE:            'PASTE',
  INSERT_ROW:       'INSERT_ROW',
  DELETE_ROW:       'DELETE_ROW',
  INSERT_COLUMN:    'INSERT_COLUMN',
  DELETE_COLUMN:    'DELETE_COLUMN',
  SHEET_CREATE:     'SHEET_CREATE',
  SHEET_DELETE:     'SHEET_DELETE',
  SHEET_RENAME:     'SHEET_RENAME',
  FORMULA_CHANGE:   'FORMULA_CHANGE',
  OTHER:            'OTHER',
};

// ============================================================
// Performance tunables
// ============================================================

/**
 * Maximum number of log rows written in a single appendRow batch.
 * Keeps execution time within Apps Script's 6-minute limit.
 */
const MAX_BATCH_SIZE = 500;

/**
 * Maximum number of cells compared in a single diff pass.
 * Ranges larger than this are chunked.
 */
const MAX_DIFF_CELLS = 10000;

// ============================================================
// Script Properties keys
// ============================================================

/** Stores the JSON array of all sheet names (for rename detection). */
const PROP_SHEET_NAMES = 'sheetNames';

/** Stores whether the audit system has been fully initialised. */
const PROP_INITIALIZED = 'auditInitialized';
