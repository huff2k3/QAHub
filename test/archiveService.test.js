const { loadGasProject } = require('./support/gasEnv');
const {
  createSpreadsheetApp,
  createPropertiesService,
  createLockService,
  createSession,
  createSheet
} = require('./support/mockGasHost');

function buildRow(CONFIG, bugId, status) {

  const row = new Array(Object.keys(CONFIG.COLUMNS).length).fill('');

  row[CONFIG.COLUMNS.BUG_ID - 1] = bugId;
  row[CONFIG.COLUMNS.STATUS - 1] = status;

  return row;

}

function setup() {

  const issuesSheet = createSheet('Issues');
  const archiveSheet = createSheet('Archive');

  const project = loadGasProject({
    SpreadsheetApp: createSpreadsheetApp({ Issues: issuesSheet, Archive: archiveSheet }),
    PropertiesService: createPropertiesService(),
    LockService: createLockService(),
    Session: createSession()
  });

  return { ...project, issuesSheet, archiveSheet };

}

describe('ArchiveService.archiveClosed', () => {

  test('returns 0 and touches nothing when Issues has no data rows', () => {

    const { ArchiveService, issuesSheet, archiveSheet } = setup();

    issuesSheet.getLastRow = jest.fn(() => 1); // header only

    expect(ArchiveService.archiveClosed()).toBe(0);
    expect(issuesSheet.deleteRow).not.toHaveBeenCalled();
    expect(archiveSheet.getRange).not.toHaveBeenCalled();

  });

  test('leaves Issues untouched when nothing is Closed', () => {

    const { ArchiveService, issuesSheet, archiveSheet, CONFIG } = setup();

    const rows = [buildRow(CONFIG, 1, 'New'), buildRow(CONFIG, 2, 'In Progress')];

    issuesSheet.getLastRow = jest.fn(() => CONFIG.FIRST_DATA_ROW + rows.length - 1);
    issuesSheet.range.getValues.mockReturnValue(rows);

    expect(ArchiveService.archiveClosed()).toBe(0);
    expect(issuesSheet.deleteRow).not.toHaveBeenCalled();
    expect(archiveSheet.range.setValues).not.toHaveBeenCalled();

  });

  test('moves only Closed rows to Archive and deletes them from Issues, bottom-up', () => {

    const { ArchiveService, issuesSheet, archiveSheet, CONFIG } = setup();

    const rows = [
      buildRow(CONFIG, 1, 'New'),               // row 2
      buildRow(CONFIG, 2, CONFIG.STATUS.CLOSED), // row 3
      buildRow(CONFIG, 3, 'In Progress'),        // row 4
      buildRow(CONFIG, 4, CONFIG.STATUS.CLOSED)  // row 5
    ];

    issuesSheet.getLastRow = jest.fn(() => CONFIG.FIRST_DATA_ROW + rows.length - 1);
    issuesSheet.range.getValues.mockReturnValue(rows);

    const archivedCount = ArchiveService.archiveClosed();

    expect(archivedCount).toBe(2);

    // Walked bottom-up: row 5 (Bug 4) first, then row 3 (Bug 2).
    expect(issuesSheet.deleteRow.mock.calls).toEqual([[5], [3]]);

    expect(archiveSheet.range.setValues.mock.calls[0][0]).toEqual([rows[3]]);
    expect(archiveSheet.range.setValues.mock.calls[1][0]).toEqual([rows[1]]);

  });

});
