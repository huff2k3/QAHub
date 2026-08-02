const { loadGasProject } = require('./support/gasEnv');
const {
  createSpreadsheetApp,
  createPropertiesService,
  createLockService,
  createSession,
  createSheet
} = require('./support/mockGasHost');

function setup() {

  const issuesSheet = createSheet('Issues', { getLastRow: jest.fn(() => 5) });
  const listsSheet = createSheet('Lists', { getLastRow: jest.fn(() => 10) });

  const project = loadGasProject({
    SpreadsheetApp: createSpreadsheetApp({ Issues: issuesSheet, Lists: listsSheet }),
    PropertiesService: createPropertiesService(),
    LockService: createLockService(),
    Session: createSession()
  });

  return { ...project, issuesSheet, listsSheet };

}

describe('ValidationService.applyToRow', () => {

  test('converts each configured column letter to its 1-based column number', () => {

    const { ValidationService, issuesSheet, CONFIG } = setup();

    ValidationService.applyToRow(issuesSheet, 7);

    const columnsUsed = issuesSheet.getRange.mock.calls
      .filter(call => call[2] === 1 && call[3] === 1)
      .map(call => call[1]);

    expect(columnsUsed).toEqual(
      expect.arrayContaining([2, 4, 3, 8]) // B, D, C, H per CONFIG.VALIDATION
    );

    expect(columnsUsed.every(col => Number.isInteger(col))).toBe(true);
    // Sanity check against CONFIG itself rather than hard-coded letters.
    Object.values(CONFIG.VALIDATION).forEach(({ issueColumn }) => {
      const expectedColumn = issueColumn.charCodeAt(0) - 64;
      expect(columnsUsed).toContain(expectedColumn);
    });

  });

  test('scopes the row range to exactly the row passed in', () => {

    const { ValidationService, issuesSheet } = setup();

    ValidationService.applyToRow(issuesSheet, 9);

    const rowCalls = issuesSheet.getRange.mock.calls.filter(call => call[2] === 1 && call[3] === 1);

    rowCalls.forEach(call => {
      expect(call[0]).toBe(9);
    });

  });

  test('reads allowed values from the matching column on the Lists sheet', () => {

    const { ValidationService, issuesSheet, listsSheet } = setup();

    ValidationService.applyToRow(issuesSheet, 2);

    expect(listsSheet.getRange).toHaveBeenCalledWith('A2:A10');
    expect(listsSheet.getRange).toHaveBeenCalledWith('B2:B10');
    expect(listsSheet.getRange).toHaveBeenCalledWith('C2:C10');
    expect(listsSheet.getRange).toHaveBeenCalledWith('D2:D10');

  });

});

describe('ValidationService.initialize', () => {

  test('applies validation across the full existing data range', () => {

    const { ValidationService, issuesSheet, CONFIG } = setup();

    ValidationService.initialize();

    const rowCalls = issuesSheet.getRange.mock.calls.filter(call => call[2] > 1 || call[3] === 1);

    expect(rowCalls.length).toBeGreaterThan(0);
    rowCalls.forEach(call => {
      expect(call[0]).toBe(CONFIG.FIRST_DATA_ROW);
    });

  });

});
