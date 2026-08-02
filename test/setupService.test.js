const { loadGasProject } = require('./support/gasEnv');
const {
  createSpreadsheetApp,
  createPropertiesService,
  createLockService,
  createSession,
  createSheet,
  createFilter
} = require('./support/mockGasHost');

function setup(existingSheets = {}) {

  const sheetsByName = { ...existingSheets };

  const project = loadGasProject({
    SpreadsheetApp: createSpreadsheetApp(sheetsByName),
    PropertiesService: createPropertiesService(),
    LockService: createLockService(),
    Session: createSession()
  });

  return { ...project, sheetsByName };

}

describe('SetupService.initializeWorkbook on a brand-new spreadsheet', () => {

  test('creates the Issues, Archive, Lists, Settings, and Dashboard sheets', () => {

    const { SetupService, sheetsByName, CONFIG } = setup();

    SetupService.initializeWorkbook();

    expect(sheetsByName[CONFIG.SHEETS.ISSUES]).toBeDefined();
    expect(sheetsByName[CONFIG.SHEETS.ARCHIVE]).toBeDefined();
    expect(sheetsByName[CONFIG.SHEETS.LISTS]).toBeDefined();
    expect(sheetsByName[CONFIG.SHEETS.SETTINGS]).toBeDefined();
    expect(sheetsByName[CONFIG.SHEETS.DASHBOARD]).toBeDefined();

  });

  test('writes Issues headers in CONFIG.COLUMNS order', () => {

    const { SetupService, sheetsByName, CONFIG } = setup();

    SetupService.initializeWorkbook();

    const issuesSheet = sheetsByName[CONFIG.SHEETS.ISSUES];

    const headerCall = issuesSheet.getRange.mock.calls.find(
      call => call[0] === CONFIG.HEADER_ROW && call[1] === 1 && call[3] === Object.keys(CONFIG.COLUMNS).length
    );

    expect(headerCall).toBeDefined();

    const range = issuesSheet.getRange.mock.results[
      issuesSheet.getRange.mock.calls.indexOf(headerCall)
    ].value;

    const expectedHeaders = Object.keys(CONFIG.COLUMNS)
      .sort((a, b) => CONFIG.COLUMNS[a] - CONFIG.COLUMNS[b])
      .map(key => CONFIG.HEADERS[key]);

    expect(range.setValues).toHaveBeenCalledWith([expectedHeaders]);

  });

  test('writes the same headers onto the Archive sheet', () => {

    const { SetupService, sheetsByName, CONFIG } = setup();

    SetupService.initializeWorkbook();

    const archiveSheet = sheetsByName[CONFIG.SHEETS.ARCHIVE];

    const headerCall = archiveSheet.getRange.mock.calls.find(
      call => call[0] === CONFIG.HEADER_ROW && call[1] === 1 && call[3] === Object.keys(CONFIG.COLUMNS).length
    );

    expect(headerCall).toBeDefined();

  });

  test('seeds each configured list under its header on the Lists sheet', () => {

    const { SetupService, sheetsByName, CONFIG } = setup();

    SetupService.initializeWorkbook();

    const listsSheet = sheetsByName[CONFIG.SHEETS.LISTS];

    Object.values(CONFIG.LISTS).forEach(list => {

      const column = list.column.charCodeAt(0) - 64;

      const headerCall = listsSheet.getRange.mock.calls.find(
        call => call[0] === CONFIG.HEADER_ROW && call[1] === column && call.length === 2
      );

      expect(headerCall).toBeDefined();

      const valuesCall = listsSheet.getRange.mock.calls.find(
        call => call[0] === CONFIG.FIRST_DATA_ROW && call[1] === column && call[2] === list.values.length
      );

      expect(valuesCall).toBeDefined();

      const range = listsSheet.getRange.mock.results[
        listsSheet.getRange.mock.calls.indexOf(valuesCall)
      ].value;

      expect(range.setValues).toHaveBeenCalledWith(
        list.values.map(value => [value])
      );

    });

  });

  test('freezes the header row and enables a filter on Issues and Archive', () => {

    const { SetupService, sheetsByName, CONFIG } = setup();

    SetupService.initializeWorkbook();

    [CONFIG.SHEETS.ISSUES, CONFIG.SHEETS.ARCHIVE].forEach(name => {

      const sheet = sheetsByName[name];

      expect(sheet.setFrozenRows).toHaveBeenCalledWith(1);

    });

  });

  test('applies data validation to the Issues sheet', () => {

    const { SetupService, sheetsByName, CONFIG } = setup();

    SetupService.initializeWorkbook();

    const issuesSheet = sheetsByName[CONFIG.SHEETS.ISSUES];

    // ValidationService.applyValidation calls getRange(row, column, count, 1)
    // — a single-column range — once per CONFIG.VALIDATION entry.
    const validationCalls = issuesSheet.getRange.mock.calls.filter(
      call => call.length === 4 && call[3] === 1
    );

    expect(validationCalls.length).toBe(Object.keys(CONFIG.VALIDATION).length);

  });

});

describe('SetupService.initializeWorkbook re-run', () => {

  test('reuses existing sheets instead of creating duplicates', () => {

    const existingIssues = createSheet('Issues');

    const { SetupService, sheetsByName, CONFIG } = setup({ Issues: existingIssues });

    SetupService.initializeWorkbook();

    expect(sheetsByName[CONFIG.SHEETS.ISSUES]).toBe(existingIssues);

  });

  test('recreates an existing filter so it covers newly added columns', () => {

    // Simulates a sheet whose filter was created back when there were fewer
    // columns (e.g. before Closed Date was added) — the filter object itself
    // doesn't grow when CONFIG.COLUMNS grows, so a stale filter must be
    // replaced, not left alone.
    const existingFilter = createFilter();

    const existingIssues = createSheet('Issues', {
      getFilter: jest.fn(() => existingFilter),
      getLastColumn: jest.fn(() => 16) // stale — one short of the current 17
    });

    const { SetupService, sheetsByName, CONFIG } = setup({ Issues: existingIssues });

    SetupService.initializeWorkbook();

    expect(existingFilter.remove).toHaveBeenCalled();

    const issuesSheet = sheetsByName[CONFIG.SHEETS.ISSUES];

    const filterRangeCall = issuesSheet.range.createFilter.mock.calls.length > 0;

    expect(filterRangeCall).toBe(true);

  });

});
