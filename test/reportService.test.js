const { loadGasProject } = require('./support/gasEnv');
const {
  createSpreadsheetApp,
  createPropertiesService,
  createLockService,
  createSession,
  createSheet,
  createDocumentApp,
  createDriveApp,
  createDriveFile,
  createDriveFolder
} = require('./support/mockGasHost');

function buildRow(CONFIG, overrides = {}) {

  const row = new Array(Object.keys(CONFIG.COLUMNS).length).fill('');

  row[CONFIG.COLUMNS.BUG_ID - 1] = overrides.bugId || '';
  row[CONFIG.COLUMNS.STATUS - 1] = overrides.status || '';
  row[CONFIG.COLUMNS.SEVERITY - 1] = overrides.severity || '';
  row[CONFIG.COLUMNS.TITLE - 1] = overrides.title || '';
  row[CONFIG.COLUMNS.ASSIGNED_TO - 1] = overrides.assignedTo || '';
  row[CONFIG.COLUMNS.NOTES - 1] = overrides.notes || '';
  row[CONFIG.COLUMNS.CREATED_DATE - 1] = overrides.createdDate || '';
  row[CONFIG.COLUMNS.LAST_UPDATED - 1] = overrides.lastUpdated || '';
  row[CONFIG.COLUMNS.CLOSED_DATE - 1] = overrides.closedDate || '';

  return row;

}

function setup(promptResponse) {

  const issuesSheet = createSheet('Issues');
  const archiveSheet = createSheet('Archive');

  const project = loadGasProject({
    SpreadsheetApp: createSpreadsheetApp({ Issues: issuesSheet, Archive: archiveSheet }, { promptResponse }),
    PropertiesService: createPropertiesService(),
    LockService: createLockService(),
    Session: createSession()
  });

  return { ...project, issuesSheet, archiveSheet };

}

describe('ReportService.isCriticalOrHigher', () => {

  test('Blocking and Critical count, High/Medium/Low do not', () => {

    const { ReportService } = setup();

    expect(ReportService.isCriticalOrHigher('Blocking')).toBe(true);
    expect(ReportService.isCriticalOrHigher('Critical')).toBe(true);
    expect(ReportService.isCriticalOrHigher('High')).toBe(false);
    expect(ReportService.isCriticalOrHigher('Medium')).toBe(false);
    expect(ReportService.isCriticalOrHigher('Low')).toBe(false);
    expect(ReportService.isCriticalOrHigher('')).toBe(false);

  });

});

describe('ReportService.buildReportData', () => {

  function daysAgo(now, n) {
    return new Date(now.getTime() - n * 24 * 60 * 60 * 1000);
  }

  test('criticalOpen only includes open, non-Closed, Critical-or-higher issues', () => {

    const { ReportService, issuesSheet, CONFIG } = setup();

    const now = new Date();

    const rows = [
      buildRow(CONFIG, { bugId: 1, status: 'New', severity: 'Critical', title: 'A' }),
      buildRow(CONFIG, { bugId: 2, status: 'In Progress', severity: 'Blocking', title: 'B' }),
      buildRow(CONFIG, { bugId: 3, status: 'New', severity: 'High', title: 'C' }),
      buildRow(CONFIG, { bugId: 4, status: CONFIG.STATUS.CLOSED, severity: 'Critical', title: 'D' })
    ];

    issuesSheet.getLastRow = jest.fn(() => CONFIG.FIRST_DATA_ROW + rows.length - 1);
    issuesSheet.range.getValues.mockReturnValue(rows);

    const data = ReportService.buildReportData(now);

    expect(data.criticalOpen.map(i => i.bugId)).toEqual([1, 2]);

  });

  test('foundAndFixed requires both Created Date and Closed Date within the lookback window', () => {

    const { ReportService, archiveSheet, CONFIG } = setup();

    const now = new Date();

    const rows = [
      // Found and fixed within the window: qualifies.
      buildRow(CONFIG, {
        bugId: 1, severity: 'Critical', status: CONFIG.STATUS.CLOSED,
        createdDate: daysAgo(now, 3), closedDate: daysAgo(now, 1)
      }),
      // Found long ago, fixed recently: created outside window, excluded.
      buildRow(CONFIG, {
        bugId: 2, severity: 'Critical', status: CONFIG.STATUS.CLOSED,
        createdDate: daysAgo(now, 30), closedDate: daysAgo(now, 1)
      }),
      // Found and fixed recently but only High severity: excluded.
      buildRow(CONFIG, {
        bugId: 3, severity: 'High', status: CONFIG.STATUS.CLOSED,
        createdDate: daysAgo(now, 2), closedDate: daysAgo(now, 1)
      })
    ];

    archiveSheet.getLastRow = jest.fn(() => CONFIG.FIRST_DATA_ROW + rows.length - 1);
    archiveSheet.range.getValues.mockReturnValue(rows);

    const data = ReportService.buildReportData(now);

    expect(data.foundAndFixed.map(i => i.bugId)).toEqual([1]);

  });

  test('notInProgressOrClosedCount excludes In Progress and Closed', () => {

    const { ReportService, issuesSheet, CONFIG } = setup();

    const rows = [
      buildRow(CONFIG, { bugId: 1, status: 'New' }),
      buildRow(CONFIG, { bugId: 2, status: 'Need More Info' }),
      buildRow(CONFIG, { bugId: 3, status: 'In Progress' }),
      buildRow(CONFIG, { bugId: 4, status: CONFIG.STATUS.CLOSED })
    ];

    issuesSheet.getLastRow = jest.fn(() => CONFIG.FIRST_DATA_ROW + rows.length - 1);
    issuesSheet.range.getValues.mockReturnValue(rows);

    const data = ReportService.buildReportData(new Date());

    expect(data.notInProgressOrClosedCount).toBe(2);

  });

  test('createdCount/closedCount only count rows within the lookback window', () => {

    const { ReportService, issuesSheet, archiveSheet, CONFIG } = setup();

    const now = new Date();

    issuesSheet.getLastRow = jest.fn(() => CONFIG.FIRST_DATA_ROW);
    issuesSheet.range.getValues.mockReturnValue([
      buildRow(CONFIG, { bugId: 1, createdDate: daysAgo(now, 2) })
    ]);

    archiveSheet.getLastRow = jest.fn(() => CONFIG.FIRST_DATA_ROW + 1);
    archiveSheet.range.getValues.mockReturnValue([
      buildRow(CONFIG, { bugId: 2, createdDate: daysAgo(now, 100), closedDate: daysAgo(now, 3) }),
      buildRow(CONFIG, { bugId: 3, createdDate: daysAgo(now, 100), closedDate: daysAgo(now, 100) })
    ]);

    const data = ReportService.buildReportData(now);

    expect(data.createdCount).toBe(1);
    expect(data.closedCount).toBe(1);

  });

});

describe('ReportService.suggestQualityStatus', () => {

  test('Red when any open Blocking/Critical issue exists', () => {

    const { ReportService, CONFIG } = setup();

    const rows = [buildRow(CONFIG, { status: 'New', severity: 'Blocking' })];

    const result = ReportService.suggestQualityStatus(rows, new Date());

    expect(result.status).toBe('Red');

  });

  test('Yellow when only High severity is open', () => {

    const { ReportService, CONFIG } = setup();

    const rows = [buildRow(CONFIG, { status: 'New', severity: 'High' })];

    const result = ReportService.suggestQualityStatus(rows, new Date());

    expect(result.status).toBe('Yellow');

  });

  test('Yellow when an open issue is stale, even at Low severity', () => {

    const { ReportService, CONFIG } = setup();

    const now = new Date();
    const staleDate = new Date(now.getTime() - (CONFIG.DASHBOARD.STALE_DAYS + 1) * 24 * 60 * 60 * 1000);

    const rows = [buildRow(CONFIG, { status: 'New', severity: 'Low', lastUpdated: staleDate })];

    const result = ReportService.suggestQualityStatus(rows, now);

    expect(result.status).toBe('Yellow');

  });

  test('Green when nothing open is Critical+/High/stale', () => {

    const { ReportService, CONFIG } = setup();

    const rows = [buildRow(CONFIG, { status: 'New', severity: 'Low', lastUpdated: new Date() })];

    const result = ReportService.suggestQualityStatus(rows, new Date());

    expect(result.status).toBe('Green');

  });

});

describe('ReportService.generate', () => {

  function setupWithDrive(promptResponse, { spreadsheetHasParent = true } = {}) {

    const base = setup(promptResponse);

    const documentApp = createDocumentApp();

    const parentFolder = spreadsheetHasParent ? createDriveFolder() : undefined;

    const driveApp = createDriveApp({
      'spreadsheet-id': createDriveFile(parentFolder),
      'doc-id': createDriveFile()
    });

    // Reload with DocumentApp/DriveApp added to the sandbox.
    const { loadGasProject } = require('./support/gasEnv');

    const project = loadGasProject({
      SpreadsheetApp: createSpreadsheetApp(
        { Issues: base.issuesSheet, Archive: base.archiveSheet },
        { promptResponse }
      ),
      PropertiesService: createPropertiesService(),
      LockService: createLockService(),
      Session: createSession(),
      DocumentApp: documentApp,
      DriveApp: driveApp
    });

    return { ...project, issuesSheet: base.issuesSheet, archiveSheet: base.archiveSheet, documentApp, driveApp, parentFolder };

  }

  test('creates a Doc, moves it into the spreadsheet\'s folder, and returns its URL', () => {

    const { ReportService, documentApp, driveApp, parentFolder } = setupWithDrive({ button: 'OK', text: 'Looking good.' });

    const url = ReportService.generate();

    expect(documentApp.create).toHaveBeenCalledTimes(1);
    expect(url).toBe('https://docs.google.com/document/d/doc-id/edit');

    expect(parentFolder.addFile).toHaveBeenCalled();
    expect(driveApp.rootFolder.removeFile).toHaveBeenCalled();

  });

  test('includes the entered comment, or "None" when the prompt is cancelled', () => {

    const { ReportService, documentApp } = setupWithDrive({ button: 'CANCEL', text: 'ignored' });

    ReportService.generate();

    const doc = documentApp.created[0];

    const commentParagraph = doc.body.paragraphs.find(p => p.text.startsWith('Comment:'));

    expect(commentParagraph.text).toBe('Comment: None');

  });

  test('does not attempt to move the Doc when the spreadsheet has no parent folder', () => {

    const { ReportService, driveApp } = setupWithDrive({ button: 'OK', text: '' }, { spreadsheetHasParent: false });

    ReportService.generate();

    expect(driveApp.rootFolder.removeFile).not.toHaveBeenCalled();

  });

});
