const { loadGasProject } = require('./support/gasEnv');
const {
  createSpreadsheetApp,
  createPropertiesService,
  createLockService,
  createSession,
  createSheet
} = require('./support/mockGasHost');

function load(overrides = {}) {

  return loadGasProject({
    SpreadsheetApp: createSpreadsheetApp(overrides.sheets || {}),
    PropertiesService: overrides.PropertiesService || createPropertiesService(),
    LockService: overrides.LockService || createLockService(),
    Session: overrides.Session || createSession()
  });

}

describe('Utils.nextBugId', () => {

  test('starts at 1 when no property has been set', () => {

    const { Utils } = load();

    expect(Utils.nextBugId()).toBe(1);

  });

  test('increments on each call and persists via PropertiesService', () => {

    const { Utils } = load();

    expect(Utils.nextBugId()).toBe(1);
    expect(Utils.nextBugId()).toBe(2);
    expect(Utils.nextBugId()).toBe(3);

  });

  test('locks and releases the script lock around the read-modify-write', () => {

    const lock = { waitLock: jest.fn(), releaseLock: jest.fn() };
    const LockService = { getScriptLock: jest.fn(() => lock) };

    const { Utils } = load({ LockService });

    Utils.nextBugId();

    expect(lock.waitLock).toHaveBeenCalledWith(30000);
    expect(lock.releaseLock).toHaveBeenCalledTimes(1);

  });

});

describe('Utils.currentUser', () => {

  test('returns the active user email when available', () => {

    const { Utils } = load({ Session: createSession('qa@example.com') });

    expect(Utils.currentUser()).toBe('qa@example.com');

  });

  test('returns an empty string when Session throws', () => {

    const Session = {
      getActiveUser: jest.fn(() => { throw new Error('no user'); })
    };

    const { Utils } = load({ Session });

    expect(Utils.currentUser()).toBe('');

  });

});

describe('Utils.getSheet / getIssuesSheet', () => {

  test('throws a descriptive error when the sheet is missing', () => {

    const { Utils } = load({ sheets: {} });

    expect(() => Utils.getIssuesSheet()).toThrow(/Issues.*not found/);

  });

  test('returns the sheet when present', () => {

    const issuesSheet = createSheet('Issues');

    const { Utils } = load({ sheets: { Issues: issuesSheet } });

    expect(Utils.getIssuesSheet()).toBe(issuesSheet);

  });

});

describe('Utils.columnLetterToNumber', () => {

  test('converts single letters to their 1-based column number', () => {

    const { Utils } = load();

    expect(Utils.columnLetterToNumber('A')).toBe(1);
    expect(Utils.columnLetterToNumber('B')).toBe(2);
    expect(Utils.columnLetterToNumber('H')).toBe(8);
    expect(Utils.columnLetterToNumber('Z')).toBe(26);

  });

});

describe('Utils.formatIssueRow', () => {

  test('applies date formatting and wraps the full row width', () => {

    const sheet = createSheet('Issues');

    const { Utils, CONFIG } = load();

    Utils.formatIssueRow(sheet, 2);

    const totalColumns = Object.keys(CONFIG.COLUMNS).length;

    expect(sheet.getRange).toHaveBeenCalledWith(2, 1, 1, totalColumns);
    expect(sheet.getRange).toHaveBeenCalledWith(2, CONFIG.COLUMNS.CREATED_DATE);
    expect(sheet.getRange).toHaveBeenCalledWith(2, CONFIG.COLUMNS.LAST_UPDATED);

  });

});
