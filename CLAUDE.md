# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

QAHub is a Google Apps Script project that turns a Google Sheet into a lightweight QA bug tracker. There is no
frontend framework or backend server — the `.gs` files are plain Apps Script (a superset of ES6 JavaScript) that
runs inside the Sheets container-bound script environment. Apps Script itself still has no local execution:
running the real thing always means pushing to a bound script and testing inside a Sheet.

## Commands

```
npm install       # installs clasp + jest (devDependencies only, nothing is published)
npm test          # runs the Jest suite (test/*.test.js) — pure Node, no Google account needed
npx jest test/utils.test.js         # run a single test file
npx jest -t "nextBugId"             # run tests matching a name

npm run login     # clasp login — one-time OAuth in a browser (needs a Google account)
npm run create    # clasp create — provisions a NEW bound Apps Script project (asks first, see below)
npm run push      # clasp push — sync local .gs files to the Apps Script project
npm run pull      # clasp pull — sync the Apps Script project back to local files
npm run open      # clasp open — opens the project in the Apps Script web editor
```

`clasp` push/pull requires a `.clasp.json` with a real `scriptId`, which in turn requires an authenticated
`clasp login` — that step can't be done headlessly. `npm test` doesn't depend on any of it.

This repo is currently linked (`.clasp.json`) to a live bound Apps Script project + Google Sheet created via `npm
run create`. `npm run push` deploys straight to it.

**Manifest note**: `appsscript.json` deliberately has no explicit `oauthScopes`. An earlier debugging session
added `"oauthScopes": ["https://www.googleapis.com/auth/spreadsheets"]` to try to fix a persistent "You do not
have permission to call SpreadsheetApp.getActiveSpreadsheet" runtime error, which made things worse (declaring
explicit scopes stops Apps Script's automatic scope detection, and the explicit list didn't cover everything the
code touches, e.g. `getUi()`). The scope declaration was reverted; the actual fix for that particular incident
was recreating the Apps Script project from scratch (its authorization state had gotten stuck). Don't re-add
`oauthScopes` unless there's a scope Apps Script genuinely can't auto-detect (e.g. an external API).

## Testing

`test/support/gasEnv.js` loads the actual `.gs` files (in the same order Apps Script would concatenate them) into
a Node `vm` sandbox seeded with fake host services (`SpreadsheetApp`, `PropertiesService`, `LockService`,
`Session` — see `test/support/mockGasHost.js`), then hands back the top-level bindings (`CONFIG`, `Utils`,
`IssueService`, etc.) for assertions. This exercises the real source files, not a copy, so the tests catch things
like the parse error `CONFIG.VALIDATION` used to have.

One gotcha specific to this setup: each `vm.createContext()` sandbox has its own realm, so `new Date()` created
inside the loaded `.gs` code is not `instanceof` the outer test file's `Date`. Compare with
`Object.prototype.toString.call(x) === '[object Date]'` instead of `toBeInstanceOf(Date)` /
`expect.any(Date)` — see `test/issueService.test.js` and `test/triggerService.test.js` for the pattern.

## Architecture

**Module pattern**: Each service file (`IssueService.gs`, `SetupService.gs`, `Utils.gs`, `ValidationService.gs`,
`ArchiveService.gs`, `SearchService.gs`) defines a single global `const` bound to an IIFE that returns a public
API object, e.g.:

```js
const IssueService = (() => {
  function createIssue() { ... }
  return { createIssue };
})();
```

Apps Script has no module system — every file shares one global scope, so these IIFEs are the only encapsulation
boundary. Keep new services in this shape and only expose what callers need via the returned object.

**Config.gs is the single source of truth** for sheet names (`CONFIG.SHEETS`), the 1-based column layout of the
Issues sheet (`CONFIG.COLUMNS`), the human-readable header text for each column (`CONFIG.HEADERS`, keyed the same
as `CONFIG.COLUMNS`), the Issues-sheet dropdown source data (`CONFIG.LISTS` — Status/Severity/Priority/
Reproducibility, each with a `column` letter that must match the corresponding `listColumn` in
`CONFIG.VALIDATION`), default field templates (`CONFIG.TEMPLATES`), script property keys, and menu labels. Any
code touching a specific column must go through `CONFIG.COLUMNS.*` rather than a hard-coded index.

**`Utils.gs` is the one canonical utility layer** — `getIssuesSheet()`/`getArchiveSheet()`/`getSheet()`,
`scriptProperties()`, `nextBugId()`, `currentUser()`, `now()`, `formatIssueRow()`, `focusTitle()`, `touch()`,
`columnLetterToNumber()`. All services (`IssueService`, `SetupService`, `ValidationService`) call through
`Utils.*` — don't reintroduce a second implementation of something that already lives there (this happened once
already with duplicate globals in Config.gs, and once more with a private `columnLetterToNumber` in
`ValidationService.gs`; both were removed in favor of the `Utils` version).

**Entry points and triggers**:
- `Code.gs` — `onOpen()` builds the "QA Tools" custom menu; menu items dispatch to `addNewIssue()`,
  `initializeWorkbook()`, `showSearch()`, and `archiveClosed()`. All four wrap their real work in try/catch and
  surface failures via `SpreadsheetApp.getUi().alert(...)` rather than letting exceptions go uncaught.
- `TriggerService.gs` — `onEdit(e)` is a simple trigger that stamps `LAST_UPDATED` whenever a cell in the
  `Issues` sheet (at or below `CONFIG.FIRST_DATA_ROW`) is edited, ignoring edits to the `LAST_UPDATED` column
  itself to avoid recursive updates.
- `SetupService.initializeWorkbook()` — builds a workbook from scratch: creates the `Issues`/`Archive`/`Lists`/
  `Settings` sheets if they don't already exist (`ensureSheet`), writes the `CONFIG.HEADERS` header row onto both
  `Issues` and `Archive`, seeds `CONFIG.LISTS` values onto `Lists`, then formats (freeze/filter/bold header/column
  widths) and applies `ValidationService.initialize()`. Re-running it is safe for issue data (never touches data
  rows) but does reset headers and Lists values back to the `CONFIG` defaults, clobbering manual edits to those
  specifically — that's intentional, not a bug.
- `ValidationService` — builds Sheets data-validation rules on the Issues sheet by reading allowed values off a
  separate `Lists` sheet (`CONFIG.VALIDATION` maps each validated Issues column letter to its source column
  letter on `Lists`, which must line up with `CONFIG.LISTS.*.column`).
- `ArchiveService.archiveClosed()` — scans Issues bottom-up for rows where `Status === CONFIG.STATUS.CLOSED`,
  appends each one onto the first empty row of `Archive`, then `deleteRow`s it from Issues. Bottom-up order is
  load-bearing: deleting a row shifts every row below it up by one, so walking top-down would skip rows.
- `SearchService.search()` — prompts for a keyword via `ui.prompt`, checks it case-insensitively against Bug ID/
  Title/Description/Keywords (`SEARCH_COLUMNS`), and jumps to the first matching row via `setActiveRange`. The
  underlying `findMatches(term)` is exposed separately and returns all matching row numbers, not just the first.

**Row shape**: A new issue is a fixed-width row built by writing directly into a sparse `values[]` array indexed
by `CONFIG.COLUMNS.* - 1` (see `IssueService.createIssue`). When adding a column, update `CONFIG.COLUMNS` —
`Utils.formatIssueRow` and `SetupService.formatHeader` both derive the column count from
`Object.keys(CONFIG.COLUMNS).length`, so they stay in sync automatically.

