const { loadGasProject } = require('./support/gasEnv');
const {
  createSpreadsheetApp,
  createPropertiesService,
  createLockService,
  createSession,
  createSheet
} = require('./support/mockGasHost');

function setup() {

  const issuesSheet = createSheet('Issues');
  const otherSheet = createSheet('Archive');

  const project = loadGasProject({
    SpreadsheetApp: createSpreadsheetApp({}),
    PropertiesService: createPropertiesService(),
    LockService: createLockService(),
    Session: createSession()
  });

  return { ...project, issuesSheet, otherSheet };

}

function makeEvent(sheet, row, column) {

  return {
    range: {
      getSheet: () => sheet,
      getRow: () => row,
      getColumn: () => column
    }
  };

}

describe('onEdit', () => {

  test('does nothing when the event object is missing (e.g. manual run)', () => {

    const { onEdit } = setup();

    expect(() => onEdit(undefined)).not.toThrow();

  });

  test('ignores edits outside the Issues sheet', () => {

    const { onEdit, otherSheet, CONFIG } = setup();

    onEdit(makeEvent(otherSheet, CONFIG.FIRST_DATA_ROW, 1));

    expect(otherSheet.getRange).not.toHaveBeenCalled();

  });

  test('ignores edits above the first data row', () => {

    const { onEdit, issuesSheet, CONFIG } = setup();

    onEdit(makeEvent(issuesSheet, CONFIG.HEADER_ROW, 1));

    expect(issuesSheet.getRange).not.toHaveBeenCalled();

  });

  test('ignores edits to the Last Updated column itself', () => {

    const { onEdit, issuesSheet, CONFIG } = setup();

    onEdit(makeEvent(issuesSheet, CONFIG.FIRST_DATA_ROW, CONFIG.COLUMNS.LAST_UPDATED));

    expect(issuesSheet.getRange).not.toHaveBeenCalled();

  });

  test('stamps Last Updated when a real field on an issue row is edited', () => {

    const { onEdit, issuesSheet, CONFIG } = setup();

    onEdit(makeEvent(issuesSheet, CONFIG.FIRST_DATA_ROW, CONFIG.COLUMNS.TITLE));

    expect(issuesSheet.getRange).toHaveBeenCalledWith(CONFIG.FIRST_DATA_ROW, CONFIG.COLUMNS.LAST_UPDATED);

    const range = issuesSheet.getRange.mock.results[0].value;

    // The sandbox has its own Date realm, so expect.any(Date) (this file's
    // Date) won't match — compare by tag instead.
    const [stamped] = range.setValue.mock.calls[0];

    expect(Object.prototype.toString.call(stamped)).toBe('[object Date]');

  });

});
