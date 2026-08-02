const { loadGasProject } = require('./support/gasEnv');
const {
  createSpreadsheetApp,
  createPropertiesService,
  createLockService,
  createSession,
  createSheet
} = require('./support/mockGasHost');

function buildRow(CONFIG, { bugId = '', title = '', description = '', keywords = '' }) {

  const row = new Array(Object.keys(CONFIG.COLUMNS).length).fill('');

  row[CONFIG.COLUMNS.BUG_ID - 1] = bugId;
  row[CONFIG.COLUMNS.TITLE - 1] = title;
  row[CONFIG.COLUMNS.DESCRIPTION - 1] = description;
  row[CONFIG.COLUMNS.KEYWORDS - 1] = keywords;

  return row;

}

function setup(promptResponse) {

  const issuesSheet = createSheet('Issues');

  const project = loadGasProject({
    SpreadsheetApp: createSpreadsheetApp({ Issues: issuesSheet }, { promptResponse }),
    PropertiesService: createPropertiesService(),
    LockService: createLockService(),
    Session: createSession()
  });

  return { ...project, issuesSheet };

}

describe('SearchService.findMatches', () => {

  test('matches case-insensitively across Bug ID, Title, Description, and Keywords', () => {

    const { SearchService, issuesSheet, CONFIG } = setup();

    const rows = [
      buildRow(CONFIG, { bugId: 1, title: 'Login crashes' }),
      buildRow(CONFIG, { bugId: 2, title: 'Fine', description: 'contains CRASH in caps' }),
      buildRow(CONFIG, { bugId: 3, title: 'Unrelated', keywords: 'crash, regression' }),
      buildRow(CONFIG, { bugId: 4, title: 'Totally fine' })
    ];

    issuesSheet.getLastRow = jest.fn(() => CONFIG.FIRST_DATA_ROW + rows.length - 1);
    issuesSheet.range.getValues.mockReturnValue(rows);

    expect(SearchService.findMatches('crash')).toEqual([2, 3, 4]);

  });

  test('returns an empty array when Issues has no data rows', () => {

    const { SearchService, issuesSheet } = setup();

    issuesSheet.getLastRow = jest.fn(() => 1);

    expect(SearchService.findMatches('anything')).toEqual([]);

  });

});

describe('SearchService.search', () => {

  test('does nothing when the user cancels the prompt', () => {

    const { SearchService, issuesSheet } = setup({ button: 'CANCEL', text: 'crash' });

    SearchService.search();

    expect(issuesSheet.getRange).not.toHaveBeenCalled();

  });

  test('does nothing when the search term is blank', () => {

    const { SearchService, issuesSheet } = setup({ button: 'OK', text: '   ' });

    SearchService.search();

    expect(issuesSheet.getRange).not.toHaveBeenCalled();

  });

  test('jumps to the first match and reports the total match count', () => {

    const { SearchService, issuesSheet, CONFIG } = setup({ button: 'OK', text: 'crash' });

    const rows = [
      buildRow(CONFIG, { bugId: 1, title: 'Fine' }),
      buildRow(CONFIG, { bugId: 2, title: 'Crash on load' }),
      buildRow(CONFIG, { bugId: 3, title: 'Also crashes' })
    ];

    issuesSheet.getLastRow = jest.fn(() => CONFIG.FIRST_DATA_ROW + rows.length - 1);
    issuesSheet.range.getValues.mockReturnValue(rows);

    SearchService.search();

    expect(issuesSheet.setActiveRange).toHaveBeenCalledWith(issuesSheet.range);
    expect(issuesSheet.getRange).toHaveBeenCalledWith(3, 1);

  });

  test('alerts when there are no matches', () => {

    const { SearchService, issuesSheet, CONFIG } = setup({ button: 'OK', text: 'nope' });

    issuesSheet.getLastRow = jest.fn(() => CONFIG.FIRST_DATA_ROW);
    issuesSheet.range.getValues.mockReturnValue([buildRow(CONFIG, { bugId: 1, title: 'Fine' })]);

    SearchService.search();

    expect(issuesSheet.setActiveRange).not.toHaveBeenCalled();

  });

});
