const { loadGasProject } = require('./support/gasEnv');
const {
  createSpreadsheetApp,
  createPropertiesService,
  createLockService,
  createSession,
  createSheet,
  createCharts
} = require('./support/mockGasHost');

function buildIssueRow(CONFIG, overrides = {}) {

  const row = new Array(Object.keys(CONFIG.COLUMNS).length).fill('');

  row[CONFIG.COLUMNS.BUG_ID - 1] = overrides.bugId || '';
  row[CONFIG.COLUMNS.STATUS - 1] = overrides.status || '';
  row[CONFIG.COLUMNS.PRIORITY - 1] = overrides.priority || '';
  row[CONFIG.COLUMNS.SEVERITY - 1] = overrides.severity || '';
  row[CONFIG.COLUMNS.TITLE - 1] = overrides.title || '';
  row[CONFIG.COLUMNS.CREATED_DATE - 1] = overrides.createdDate || '';
  row[CONFIG.COLUMNS.LAST_UPDATED - 1] = overrides.lastUpdated || '';

  return row;

}

function setup() {

  const dashboardSheet = createSheet('Dashboard');
  const issuesSheet = createSheet('Issues');
  const archiveSheet = createSheet('Archive');

  const sheetsByName = {
    Dashboard: dashboardSheet,
    Issues: issuesSheet,
    Archive: archiveSheet
  };

  const project = loadGasProject({
    SpreadsheetApp: createSpreadsheetApp(sheetsByName),
    PropertiesService: createPropertiesService(),
    LockService: createLockService(),
    Session: createSession(),
    Charts: createCharts()
  });

  return { ...project, dashboardSheet, issuesSheet, archiveSheet };

}

describe('DashboardService.refresh', () => {

  test('creates the Dashboard sheet if it does not exist yet', () => {

    const issuesSheet = createSheet('Issues');
    const archiveSheet = createSheet('Archive');
    const sheetsByName = { Issues: issuesSheet, Archive: archiveSheet };

    const { DashboardService } = loadGasProject({
      SpreadsheetApp: createSpreadsheetApp(sheetsByName),
      PropertiesService: createPropertiesService(),
      LockService: createLockService(),
      Session: createSession(),
      Charts: createCharts()
    });

    DashboardService.refresh();

    expect(sheetsByName.Dashboard).toBeDefined();

  });

  test('clears old content and removes existing charts before rebuilding', () => {

    const oldChart = { __old: true };

    const { DashboardService, dashboardSheet } = setup();

    dashboardSheet.getCharts = jest.fn(() => [oldChart]);

    DashboardService.refresh();

    expect(dashboardSheet.clear).toHaveBeenCalled();
    expect(dashboardSheet.removeChart).toHaveBeenCalledWith(oldChart);

  });

  test('counts issues by Status, including unset values as "(Not Set)"', () => {

    const { DashboardService, issuesSheet, dashboardSheet, CONFIG } = setup();

    const rows = [
      buildIssueRow(CONFIG, { bugId: 1, status: 'New' }),
      buildIssueRow(CONFIG, { bugId: 2, status: 'New' }),
      buildIssueRow(CONFIG, { bugId: 3, status: 'Closed' }),
      buildIssueRow(CONFIG, { bugId: 4, status: '' })
    ];

    issuesSheet.getLastRow = jest.fn(() => CONFIG.FIRST_DATA_ROW + rows.length - 1);
    issuesSheet.range.getValues.mockReturnValue(rows);

    DashboardService.refresh();

    const statusTableCall = dashboardSheet.range.setValues.mock.calls.find(
      call => Array.isArray(call[0]) && call[0].some(r => r[0] === 'New')
    );

    expect(statusTableCall).toBeDefined();

    const table = statusTableCall[0];

    expect(table.find(r => r[0] === 'New')[1]).toBe(2);
    expect(table.find(r => r[0] === 'Closed')[1]).toBe(1);
    expect(table.find(r => r[0] === '(Not Set)')[1]).toBe(1);

  });

  test('inserts a chart for each breakdown section plus throughput', () => {

    const { DashboardService, dashboardSheet } = setup();

    DashboardService.refresh();

    // Status, Priority, Severity, Throughput = 4 charts.
    expect(dashboardSheet.insertChart).toHaveBeenCalledTimes(4);

  });

  test('lists stale open issues sorted oldest-first, excluding Closed issues', () => {

    const { DashboardService, issuesSheet, dashboardSheet, CONFIG } = setup();

    const now = new Date();

    const daysAgo = n => new Date(now.getTime() - n * 24 * 60 * 60 * 1000);

    const rows = [
      buildIssueRow(CONFIG, { bugId: 1, status: 'New', title: 'Fresh', lastUpdated: daysAgo(1) }),
      buildIssueRow(CONFIG, { bugId: 2, status: 'In Progress', title: 'Old', lastUpdated: daysAgo(30) }),
      buildIssueRow(CONFIG, { bugId: 3, status: 'In Progress', title: 'Oldest', lastUpdated: daysAgo(90) }),
      buildIssueRow(CONFIG, { bugId: 4, status: CONFIG.STATUS.CLOSED, title: 'Closed but old', lastUpdated: daysAgo(90) })
    ];

    issuesSheet.getLastRow = jest.fn(() => CONFIG.FIRST_DATA_ROW + rows.length - 1);
    issuesSheet.range.getValues.mockReturnValue(rows);

    DashboardService.refresh();

    const staleTableCall = dashboardSheet.range.setValues.mock.calls.find(
      call => Array.isArray(call[0]) && call[0].some(r => r[1] === 'Oldest')
    );

    expect(staleTableCall).toBeDefined();

    const staleTitles = staleTableCall[0].map(r => r[1]);

    expect(staleTitles).toEqual(['Oldest', 'Old']);

  });

});
