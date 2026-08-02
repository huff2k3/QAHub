const { loadGasProject } = require('./support/gasEnv');
const {
  createSpreadsheetApp,
  createPropertiesService,
  createLockService,
  createSession,
  createSheet
} = require('./support/mockGasHost');

describe('IssueService.createIssue', () => {

  function setup() {

    const issuesSheet = createSheet('Issues');
    const listsSheet = createSheet('Lists');

    const project = loadGasProject({
      SpreadsheetApp: createSpreadsheetApp({ Issues: issuesSheet, Lists: listsSheet }),
      PropertiesService: createPropertiesService(),
      LockService: createLockService(),
      Session: createSession('creator@example.com')
    });

    return { ...project, issuesSheet, listsSheet };

  }

  test('inserts a new row directly under the header', () => {

    const { IssueService, issuesSheet, CONFIG } = setup();

    IssueService.createIssue();

    expect(issuesSheet.insertRowBefore).toHaveBeenCalledWith(CONFIG.FIRST_DATA_ROW);

  });

  test('writes a full row of default values at the expected columns', () => {

    const { IssueService, issuesSheet, CONFIG } = setup();

    IssueService.createIssue();

    const totalColumns = Object.keys(CONFIG.COLUMNS).length;

    const writeCall = issuesSheet.getRange.mock.calls.find(
      call => call[0] === CONFIG.FIRST_DATA_ROW && call[1] === 1 && call[3] === totalColumns
    );

    expect(writeCall).toBeDefined();

    const range = issuesSheet.getRange.mock.results[0].value;

    const [values] = range.setValues.mock.calls[0][0];

    expect(values[CONFIG.COLUMNS.BUG_ID - 1]).toBe(1);
    expect(values[CONFIG.COLUMNS.STATUS - 1]).toBe(CONFIG.STATUS.DEFAULT);
    expect(values[CONFIG.COLUMNS.CREATOR - 1]).toBe('creator@example.com');
    expect(values[CONFIG.COLUMNS.DESCRIPTION - 1]).toBe(CONFIG.TEMPLATES.DESCRIPTION);
    expect(values[CONFIG.COLUMNS.STEPS - 1]).toBe(CONFIG.TEMPLATES.STEPS);
    // The sandbox has its own Date realm, so `new Date()` there isn't
    // `instanceof` this file's Date — compare by tag instead.
    expect(Object.prototype.toString.call(values[CONFIG.COLUMNS.CREATED_DATE - 1])).toBe('[object Date]');
    expect(values[CONFIG.COLUMNS.LAST_UPDATED - 1]).toBe(values[CONFIG.COLUMNS.CREATED_DATE - 1]);

  });

  test('assigns sequential bug IDs across multiple issues', () => {

    const { IssueService, issuesSheet, CONFIG } = setup();

    IssueService.createIssue();
    IssueService.createIssue();

    const sharedRange = issuesSheet.getRange.mock.results[0].value;

    const firstBugId = sharedRange.setValues.mock.calls[0][0][0][CONFIG.COLUMNS.BUG_ID - 1];
    const secondBugId = sharedRange.setValues.mock.calls[1][0][0][CONFIG.COLUMNS.BUG_ID - 1];

    expect(secondBugId).toBe(firstBugId + 1);

  });

  test('applies dropdown validation to the new row for every configured column', () => {

    const { IssueService, issuesSheet, CONFIG } = setup();

    IssueService.createIssue();

    const validationCalls = issuesSheet.getRange.mock.calls.filter(
      call => call[0] === CONFIG.FIRST_DATA_ROW && call.length === 4 && call[3] === 1
    );

    const validatedColumns = validationCalls.map(call => call[1]).sort((a, b) => a - b);

    const expectedColumns = Object.values(CONFIG.VALIDATION)
      .map(v => v.issueColumn.charCodeAt(0) - 64)
      .sort((a, b) => a - b);

    expect(validatedColumns).toEqual(expectedColumns);

  });

  test('moves the cursor to the Title cell after creating the row', () => {

    const { IssueService, issuesSheet, CONFIG } = setup();

    IssueService.createIssue();

    expect(issuesSheet.setActiveSelection).toHaveBeenCalled();
    expect(issuesSheet.getRange).toHaveBeenCalledWith(CONFIG.FIRST_DATA_ROW, CONFIG.COLUMNS.TITLE);

  });

  test('throws when the Issues sheet does not exist', () => {

    const project = loadGasProject({
      SpreadsheetApp: createSpreadsheetApp({}),
      PropertiesService: createPropertiesService(),
      LockService: createLockService(),
      Session: createSession()
    });

    expect(() => project.IssueService.createIssue()).toThrow(/Issues.*not found/);

  });

});
