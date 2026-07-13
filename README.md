# 🔍 Google Sheets Audit Trail

A production-ready, zero-dependency audit logging system for Google Sheets built entirely with **Google Apps Script**. Every meaningful change is captured with full before/after context — who changed what, when, and from which value.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Platform: Google Apps Script](https://img.shields.io/badge/Platform-Google%20Apps%20Script-brightgreen)](https://developers.google.com/apps-script)

---

## ✨ Features

| Capability | Details |
|---|---|
| **Single-cell edits** | Logs old value → new value for every edit |
| **Multi-cell pastes** | Per-cell diffing via snapshot cache — only changed cells are logged |
| **Formula changes** | Detects formula-text changes; ignores computed-value-only changes |
| **Row/Column mutations** | INSERT_ROW, DELETE_ROW, INSERT_COLUMN, DELETE_COLUMN |
| **Sheet operations** | SHEET_CREATE, SHEET_DELETE, SHEET_RENAME |
| **Formatting changes** | Explicitly **ignored** per spec |
| **Recursive logging** | Fully prevented — Audit Log and snapshot sheets are always skipped |
| **Performance** | Diff is scoped to the affected range only; batch writes via setValues() |

---

## 🏗️ Architecture

```
google-sheets-audit-trail/
├── src/
│   ├── Config.js      # Constants, column headers, change-type enums
│   ├── Logger.js      # Audit Log sheet management & batch write
│   ├── Snapshot.js    # Hidden cache sheets for previous-state tracking
│   ├── Diff.js        # Per-cell change detection (values + formulas)
│   ├── Triggers.js    # Event handlers: onEdit, onChange, onOpen
│   └── Setup.js       # One-time initialization & trigger installation
├── appsscript.json    # Apps Script project manifest
└── README.md
```

### Data Flow

```
User makes a change
        │
        ▼
  onEditTrigger(e)   ◄──────── onChange fires for structural ops
        │
        ▼
  Is sheet ignored? ──YES──► return
        │ NO
        ▼
  Single cell?
   ├── YES: use e.oldValue / e.value + formula check (Diff.js)
   └── NO:  compare range vs Snapshot (Diff.js → detectCellChanges)
        │
        ▼
  Build log rows (Logger.js → buildLogRow)
        │
        ▼
  Batch-write to "Audit Log" sheet (Logger.js → logEntries)
        │
        ▼
  Update Snapshot for edited range (Snapshot.js → updateSnapshotRange)
```

### Snapshot Architecture

For every tracked sheet named e.g. `Sales`, two hidden companion sheets are maintained:

| Hidden Sheet | Contents |
|---|---|
| `__snapshot__Sales` | Last-known **cell values** (getValues() output) |
| `__snapformula__Sales` | Last-known **cell formulas** (getFormulas() output) |

These enable the system to always know the *previous* state of any cell, even when Apps Script's event object doesn't provide `e.oldValue` (e.g. for pastes or multi-cell edits).

---

## 📋 Audit Log Format

Each entry in the **Audit Log** sheet contains:

| Column | Description | Example |
|---|---|---|
| Timestamp | ISO date/time of the change | `2024-01-15 09:30:00` |
| User | Email of the acting user | `alice@company.com` |
| Spreadsheet Name | Name of the spreadsheet | `Q1 Budget` |
| Spreadsheet ID | Unique spreadsheet ID | `1BxiMVs0…` |
| Sheet | Tab name where change occurred | `Sales` |
| Cell | A1 notation of the cell | `B5` |
| Row | 1-indexed row number | `5` |
| Column | 1-indexed column number | `2` |
| Old Value | Previous value (or formula) | `1500` |
| New Value | New value (or formula) | `2000` |
| Change Type | Category of change | `EDIT` |

### Change Types

| Type | Description |
|---|---|
| `EDIT` | Single-cell or multi-cell plain value change |
| `PASTE` | Multi-cell paste (one row per changed cell) |
| `FORMULA_CHANGE` | A cell's formula text was modified |
| `INSERT_ROW` | One or more rows were inserted |
| `DELETE_ROW` | One or more rows were deleted |
| `INSERT_COLUMN` | One or more columns were inserted |
| `DELETE_COLUMN` | One or more columns were deleted |
| `SHEET_CREATE` | A new sheet tab was created |
| `SHEET_DELETE` | A sheet tab was deleted |
| `SHEET_RENAME` | A sheet tab was renamed (Old Value = old name, New Value = new name) |
| `OTHER` | Unclassified structural change |

---

## 🚀 Installation

### Prerequisites

- A Google Account
- Access to the target Google Spreadsheet
- Google Apps Script editor access (Extensions → Apps Script)

### Step 1 — Copy the source files

Open your Google Spreadsheet and go to **Extensions → Apps Script**.

Create the following script files and paste in the corresponding source code:

| File to create in Apps Script | Source file in this repo |
|---|---|
| `Config.gs` | `src/Config.js` |
| `Logger.gs` | `src/Logger.js` |
| `Snapshot.gs` | `src/Snapshot.js` |
| `Diff.gs` | `src/Diff.js` |
| `Triggers.gs` | `src/Triggers.js` |
| `Setup.gs` | `src/Setup.js` |

> **Tip**: The `.gs` extension used in the Apps Script editor is equivalent to `.js` — they are the same language.

### Step 2 — Set the manifest

In the Apps Script editor, click the **Project Settings** gear (⚙️), enable "Show `appsscript.json`", then replace the contents of `appsscript.json` with the file from this repo.

### Step 3 — Run the initializer

In the Apps Script editor:

1. Select the function `initializeAuditSystem` from the function dropdown.
2. Click **▶ Run**.
3. Google will prompt for the required permissions — click **Allow**.
4. Return to your spreadsheet — you'll see an **"Audit Trail Active"** confirmation dialog.
5. A new **Audit Log** sheet will appear, and a **🔍 Audit Trail** menu will be added.

### Step 4 — Verify triggers

Go to **Extensions → Apps Script → Triggers** (clock icon ⏱️ in the left sidebar) and confirm three triggers exist:

| Function | Event source | Event type |
|---|---|---|
| `onEditTrigger` | From spreadsheet | On edit |
| `onChangeTrigger` | From spreadsheet | On change |
| `onOpenTrigger` | From spreadsheet | On open |

---

## 🔐 Required Permissions

When you run `initializeAuditSystem()`, Google will ask you to grant:

| Scope | Reason |
|---|---|
| `spreadsheets` | Read/write cells, create sheets, manage snapshots |
| `script.scriptapp` | Install installable triggers |
| `userinfo.email` | Capture the acting user's email in log entries |

> **Note**: If you are not the spreadsheet owner, trigger creation requires that you grant these permissions under **your own** Google account. Each user who needs triggers must run `initializeAuditSystem()` at least once.

---

## ⚡ Performance

| Scenario | Strategy |
|---|---|
| Single-cell edit | Uses event object values directly — no sheet read |
| Multi-cell paste | Reads only the pasted range, diffs against snapshot slice |
| Snapshot update | Writes only the affected range, not the whole sheet |
| Log writes | Single `setValues()` call for all rows — no per-row appendRow |
| Large pastes (>10k cells) | Chunked diff passes to stay within memory limits |

Tested to work efficiently on sheets with **10,000+ rows**.

---

## ⚠️ Known Limitations

1. **User email in simple trigger context**: Apps Script restricts `Session.getActiveUser().getEmail()` in simple triggers. This system uses installable triggers, which do have email access — but only if the trigger was installed by the same user. If multiple users share a spreadsheet, each user should run `initializeAuditSystem()` to register triggers under their own account.

2. **Sheet renames**: Apps Script's `onChange` event does not directly provide the old and new sheet names. The rename is detected by comparing the current sheet list against a list stored in ScriptProperties. In the rare case of a simultaneous rename + create/delete, the rename may not be detected; it will fall through as `OTHER`.

3. **Undo operations**: If a user presses Ctrl+Z, Apps Script fires another `onEdit` event for the undo, which will be logged as a new change. This is by design and provides a full chronological trail including undos.

4. **Script execution time**: Apps Script has a 6-minute execution limit per run. Very large pastes (hundreds of thousands of cells) may exceed this. The `MAX_DIFF_CELLS` constant in `Config.js` can be lowered to mitigate this.

5. **Formula-only dependency changes**: If cell `B1 = A1 + 1` and you change `A1`, the computed value of `B1` changes but its formula does not. Per spec, this is **not logged**. Only the edit to `A1` itself is logged.

6. **Apps Script quota**: Google limits the number of trigger executions and spreadsheet API calls per day. For very high-traffic spreadsheets (thousands of edits/day), consider setting up a time-based trigger to batch-flush logs instead of logging in real time.

---

## 🔧 Configuration

All tuneable constants are in `src/Config.js`:

```javascript
// Change the audit log sheet name
const AUDIT_LOG_SHEET_NAME = 'Audit Log';

// Adjust performance limits
const MAX_BATCH_SIZE  = 500;   // Max rows written per batch
const MAX_DIFF_CELLS  = 10000; // Max cells compared per diff pass
```

---

## 🛠️ Development & Deployment

### Recommended tooling: clasp

[clasp](https://github.com/google/clasp) is Google's official CLI for Apps Script development. It allows you to edit files locally and push to Apps Script.

```bash
# Install clasp
npm install -g @google/clasp

# Authenticate
clasp login

# Clone an existing Apps Script project
clasp clone <scriptId>

# Push local changes to Apps Script
clasp push

# Pull remote changes
clasp pull
```

> **Important**: Never commit `.clasprc.json` or `.clasp.json` to version control — they contain your OAuth tokens and project IDs.

---

## 📄 License

MIT © 2024

---

## 🤝 Contributing

Contributions are welcome! Please open an issue before submitting a pull request so we can discuss the proposed change.
