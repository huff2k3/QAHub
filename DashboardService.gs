/*
 * QAHub
 * Version: 1.0.0
 *
 * File:
 * DashboardService.gs
 */

const DashboardService = (() => {

  const MS_PER_DAY = 24 * 60 * 60 * 1000;

  // Tables stack in columns A-B; charts stack independently in column E so a
  // tall chart never overlaps the next table.
  const TABLE_COLUMN = 1;
  const CHART_COLUMN = 5;
  const SECTION_GAP = 2;
  const CHART_ROW_SPAN = 18;

  /**
   * Rebuilds the Dashboard sheet from the current Issues and Archive data:
   * status/priority/severity breakdowns, weekly created-vs-closed
   * throughput, and a list of stale (long-untouched) open issues.
   */
  function refresh() {

    const sheet = Utils.ensureSheet(CONFIG.SHEETS.DASHBOARD);

    sheet.clear();
    sheet.getCharts().forEach(chart => sheet.removeChart(chart));

    const openRows = Utils.readDataRows(Utils.getIssuesSheet());
    const closedRows = Utils.readDataRows(Utils.getArchiveSheet());

    let tableRow = writeTitle(sheet);
    let chartRow = tableRow;

    ({ tableRow, chartRow } = writeBreakdownSection(
      sheet, tableRow, chartRow, 'Issues by Status',
      CONFIG.COLUMNS.STATUS, CONFIG.LISTS.STATUS.values, openRows
    ));

    ({ tableRow, chartRow } = writeBreakdownSection(
      sheet, tableRow, chartRow, 'Issues by Priority',
      CONFIG.COLUMNS.PRIORITY, CONFIG.LISTS.PRIORITY.values, openRows
    ));

    ({ tableRow, chartRow } = writeBreakdownSection(
      sheet, tableRow, chartRow, 'Issues by Severity',
      CONFIG.COLUMNS.SEVERITY, CONFIG.LISTS.SEVERITY.values, openRows
    ));

    ({ tableRow, chartRow } = writeThroughputSection(sheet, tableRow, chartRow, openRows, closedRows));

    writeStaleSection(sheet, tableRow, openRows);

  }

  /**
   * Writes the dashboard title + refresh timestamp, returns the next free row.
   */
  function writeTitle(sheet) {

    sheet.getRange(1, TABLE_COLUMN)
      .setValue('QAHub Dashboard')
      .setFontWeight('bold')
      .setFontSize(14);

    sheet.getRange(2, TABLE_COLUMN)
      .setValue(`Last refreshed: ${formatDateLabel(Utils.now())}`);

    return 4;

  }

  /**
   * Writes a "count of open issues by <column>" table + column chart.
   *
   * @returns {{tableRow:number, chartRow:number}} next free row in each track.
   */
  function writeBreakdownSection(sheet, tableRow, chartRow, title, columnIndex, knownValues, rows) {

    const entries = countByColumn(rows, columnIndex, knownValues);

    sheet.getRange(tableRow, TABLE_COLUMN).setValue(title).setFontWeight('bold');

    const headerRow = tableRow + 1;

    sheet.getRange(headerRow, TABLE_COLUMN, 1, 2).setValues([['Value', 'Count']]).setFontWeight('bold');

    sheet.getRange(headerRow + 1, TABLE_COLUMN, entries.length, 2).setValues(entries);

    const dataRange = sheet.getRange(headerRow, TABLE_COLUMN, entries.length + 1, 2);

    insertColumnChart(sheet, dataRange, chartRow, CHART_COLUMN, title);

    return {
      tableRow: headerRow + 1 + entries.length + SECTION_GAP,
      chartRow: chartRow + CHART_ROW_SPAN
    };

  }

  /**
   * Counts rows per known category value (in knownValues order), then any
   * unrecognized values, then a "(Not Set)" bucket for blanks.
   *
   * @returns {Array<[string, number]>}
   */
  function countByColumn(rows, columnIndex, knownValues) {

    const counts = {};

    knownValues.forEach(value => { counts[value] = 0; });

    let unspecified = 0;

    rows.forEach(row => {

      const value = row[columnIndex - 1];

      if (!value) {
        unspecified++;
        return;
      }

      counts[value] = (counts[value] || 0) + 1;

    });

    const entries = knownValues.map(value => [value, counts[value]]);

    Object.keys(counts)
      .filter(value => knownValues.indexOf(value) === -1)
      .sort()
      .forEach(value => entries.push([value, counts[value]]));

    if (unspecified > 0) {
      entries.push(['(Not Set)', unspecified]);
    }

    return entries;

  }

  /**
   * Writes a weekly "issues created vs. issues closed" table + chart.
   * "Closed" is approximated as an archived issue's Last Updated date,
   * since QAHub doesn't track a separate closed-date field.
   */
  function writeThroughputSection(sheet, tableRow, chartRow, openRows, closedRows) {

    const title = 'Weekly Throughput (Created vs. Closed)';

    const created = countByWeek(openRows.concat(closedRows), CONFIG.COLUMNS.CREATED_DATE);
    const closed = countByWeek(closedRows, CONFIG.COLUMNS.LAST_UPDATED);

    const weeks = Array.from(
      new Set([...Object.keys(created), ...Object.keys(closed)])
    ).sort();

    const entries = weeks.map(week => [week, created[week] || 0, closed[week] || 0]);

    sheet.getRange(tableRow, TABLE_COLUMN).setValue(title).setFontWeight('bold');

    const headerRow = tableRow + 1;

    sheet.getRange(headerRow, TABLE_COLUMN, 1, 3)
      .setValues([['Week Of', 'Created', 'Closed']])
      .setFontWeight('bold');

    if (entries.length > 0) {
      sheet.getRange(headerRow + 1, TABLE_COLUMN, entries.length, 3).setValues(entries);
    }

    const dataRange = sheet.getRange(headerRow, TABLE_COLUMN, entries.length + 1, 3);

    insertColumnChart(sheet, dataRange, chartRow, CHART_COLUMN, title);

    return {
      tableRow: headerRow + 1 + entries.length + SECTION_GAP,
      chartRow: chartRow + CHART_ROW_SPAN
    };

  }

  /**
   * Buckets rows into { "yyyy-MM-dd" (start of week): count } by a date column.
   */
  function countByWeek(rows, dateColumnIndex) {

    const counts = {};

    rows.forEach(row => {

      const date = row[dateColumnIndex - 1];

      if (!date) {
        return;
      }

      const week = weekLabel(date);

      counts[week] = (counts[week] || 0) + 1;

    });

    return counts;

  }

  /**
   * Writes a table of open (non-Closed) issues whose Last Updated is at
   * least CONFIG.DASHBOARD.STALE_DAYS old, oldest first.
   */
  function writeStaleSection(sheet, tableRow, openRows) {

    const now = Utils.now();

    const stale = openRows
      .filter(row => row[CONFIG.COLUMNS.STATUS - 1] !== CONFIG.STATUS.CLOSED)
      .map(row => ({
        bugId: row[CONFIG.COLUMNS.BUG_ID - 1],
        title: row[CONFIG.COLUMNS.TITLE - 1],
        status: row[CONFIG.COLUMNS.STATUS - 1],
        lastUpdated: row[CONFIG.COLUMNS.LAST_UPDATED - 1],
        daysSinceUpdate: row[CONFIG.COLUMNS.LAST_UPDATED - 1]
          ? Math.floor((now - row[CONFIG.COLUMNS.LAST_UPDATED - 1]) / MS_PER_DAY)
          : null
      }))
      .filter(issue => issue.daysSinceUpdate !== null && issue.daysSinceUpdate >= CONFIG.DASHBOARD.STALE_DAYS)
      .sort((a, b) => b.daysSinceUpdate - a.daysSinceUpdate);

    const title = `Stale Issues (no update in ${CONFIG.DASHBOARD.STALE_DAYS}+ days)`;

    sheet.getRange(tableRow, TABLE_COLUMN).setValue(title).setFontWeight('bold');

    const headerRow = tableRow + 1;

    sheet.getRange(headerRow, TABLE_COLUMN, 1, 4)
      .setValues([['Bug ID', 'Title', 'Status', 'Days Since Update']])
      .setFontWeight('bold');

    if (stale.length === 0) {
      sheet.getRange(headerRow + 1, TABLE_COLUMN).setValue('None — nice work.');
      return headerRow + 2;
    }

    const entries = stale.map(issue => [issue.bugId, issue.title, issue.status, issue.daysSinceUpdate]);

    sheet.getRange(headerRow + 1, TABLE_COLUMN, entries.length, 4).setValues(entries);

    return headerRow + 1 + entries.length;

  }

  /**
   * Inserts a column chart over a data range, replacing whatever chart (if
   * any) refresh() already cleared from this sheet.
   */
  function insertColumnChart(sheet, dataRange, anchorRow, anchorColumn, title) {

    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataRange)
      .setPosition(anchorRow, anchorColumn, 0, 0)
      .setOption('title', title)
      .build();

    sheet.insertChart(chart);

  }

  function startOfWeek(date) {

    const copy = new Date(date.getFullYear(), date.getMonth(), date.getDate());

    copy.setDate(copy.getDate() - copy.getDay());

    return copy;

  }

  function formatDateLabel(date) {

    const pad = n => String(n).padStart(2, '0');

    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;

  }

  function weekLabel(date) {
    return formatDateLabel(startOfWeek(date));
  }

  return {
    refresh
  };

})();
