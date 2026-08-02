/*
 * QAHub
 * Version: 1.0.0
 *
 * File:
 * ReportService.gs
 */

const ReportService = (() => {

  /**
   * Builds report data, prompts for an optional quality-status comment, and
   * writes it all to a new Google Doc placed alongside the spreadsheet.
   *
   * @returns {string} The created Doc's URL.
   */
  function generate() {

    const now = Utils.now();

    const data = buildReportData(now);

    const comment = promptForComment(data);

    const doc = buildDocument(data, comment, now);

    return doc.getUrl();

  }

  /**
   * Computes every number/list the report needs, from the current Issues
   * and Archive data. Pure aside from reading the sheets — no UI, no Doc.
   *
   * @param {Date} now
   */
  function buildReportData(now) {

    const openRows = Utils.readDataRows(Utils.getIssuesSheet());
    const closedRows = Utils.readDataRows(Utils.getArchiveSheet());
    const allRows = openRows.concat(closedRows);

    const criticalOpen = openRows
      .filter(row => row[CONFIG.COLUMNS.STATUS - 1] !== CONFIG.STATUS.CLOSED)
      .filter(row => isCriticalOrHigher(row[CONFIG.COLUMNS.SEVERITY - 1]))
      .map(row => ({
        bugId: row[CONFIG.COLUMNS.BUG_ID - 1],
        title: row[CONFIG.COLUMNS.TITLE - 1],
        status: row[CONFIG.COLUMNS.STATUS - 1],
        severity: row[CONFIG.COLUMNS.SEVERITY - 1],
        assignedTo: row[CONFIG.COLUMNS.ASSIGNED_TO - 1],
        notes: row[CONFIG.COLUMNS.NOTES - 1]
      }));

    const foundAndFixed = allRows
      .filter(row => isCriticalOrHigher(row[CONFIG.COLUMNS.SEVERITY - 1]))
      .filter(row => {

        const created = row[CONFIG.COLUMNS.CREATED_DATE - 1];
        const closed = row[CONFIG.COLUMNS.CLOSED_DATE - 1];

        return withinLookback(created, now) && closed && withinLookback(closed, now);

      })
      .map(row => ({
        bugId: row[CONFIG.COLUMNS.BUG_ID - 1],
        title: row[CONFIG.COLUMNS.TITLE - 1],
        severity: row[CONFIG.COLUMNS.SEVERITY - 1],
        createdDate: row[CONFIG.COLUMNS.CREATED_DATE - 1],
        closedDate: row[CONFIG.COLUMNS.CLOSED_DATE - 1]
      }));

    const notInProgressOrClosedCount = openRows.filter(row => {

      const status = row[CONFIG.COLUMNS.STATUS - 1];

      return status !== 'In Progress' && status !== CONFIG.STATUS.CLOSED;

    }).length;

    const createdCount = allRows.filter(
      row => withinLookback(row[CONFIG.COLUMNS.CREATED_DATE - 1], now)
    ).length;

    const closedCount = allRows.filter(row => {

      const closed = row[CONFIG.COLUMNS.CLOSED_DATE - 1];

      return closed && withinLookback(closed, now);

    }).length;

    const quality = suggestQualityStatus(openRows, now);

    return {
      generatedAt: now,
      criticalOpen,
      foundAndFixed,
      notInProgressOrClosedCount,
      createdCount,
      closedCount,
      qualityStatus: quality.status,
      qualityReason: quality.reason
    };

  }

  /**
   * Suggests an overall Red/Yellow/Green quality status from simple
   * thresholds — a starting point, not a verdict; generate() lets the
   * person running the report add a comment alongside it.
   */
  function suggestQualityStatus(openRows, now) {

    const openNonClosed = openRows.filter(row => row[CONFIG.COLUMNS.STATUS - 1] !== CONFIG.STATUS.CLOSED);

    const hasCriticalOrHigher = openNonClosed.some(
      row => isCriticalOrHigher(row[CONFIG.COLUMNS.SEVERITY - 1])
    );

    if (hasCriticalOrHigher) {
      return { status: 'Red', reason: 'One or more Blocking/Critical issues are currently open.' };
    }

    const hasHigh = openNonClosed.some(row => row[CONFIG.COLUMNS.SEVERITY - 1] === 'High');

    const hasStale = openNonClosed.some(row => {

      const lastUpdated = row[CONFIG.COLUMNS.LAST_UPDATED - 1];

      return lastUpdated && Utils.daysSince(lastUpdated, now) >= CONFIG.DASHBOARD.STALE_DAYS;

    });

    if (hasHigh) {
      return { status: 'Yellow', reason: 'One or more High severity issues are open.' };
    }

    if (hasStale) {
      return {
        status: 'Yellow',
        reason: `One or more open issues haven't been updated in ${CONFIG.DASHBOARD.STALE_DAYS}+ days.`
      };
    }

    return { status: 'Green', reason: 'No Blocking/Critical/High issues open, nothing stale.' };

  }

  /**
   * Whether a severity is at or above CONFIG.REPORT.CRITICAL_THRESHOLD in
   * CONFIG.LISTS.SEVERITY.values (which is ordered most-to-least severe).
   */
  function isCriticalOrHigher(severity) {

    const order = CONFIG.LISTS.SEVERITY.values;

    const thresholdIndex = order.indexOf(CONFIG.REPORT.CRITICAL_THRESHOLD);
    const severityIndex = order.indexOf(severity);

    return severityIndex !== -1 && severityIndex <= thresholdIndex;

  }

  /**
   * Whether a date falls within the last CONFIG.REPORT.LOOKBACK_DAYS days.
   */
  function withinLookback(date, now) {
    return !!date && Utils.daysSince(date, now) <= CONFIG.REPORT.LOOKBACK_DAYS;
  }

  /**
   * Shows the suggested quality status and asks for an optional comment.
   * Returns '' if the user cancels or leaves it blank.
   */
  function promptForComment(data) {

    const ui = SpreadsheetApp.getUi();

    const response = ui.prompt(
      'Quality Status',
      `Suggested status: ${data.qualityStatus} — ${data.qualityReason}\n\n` +
      'Add an optional comment to include in the report (or leave blank):',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
      return '';
    }

    return response.getResponseText().trim();

  }

  /**
   * Writes the report into a new Google Doc placed in the same Drive folder
   * as the spreadsheet, and returns the Doc.
   */
  function buildDocument(data, comment, now) {

    const doc = DocumentApp.create(`QAHub Status Report - ${Utils.formatDateLabel(now)}`);

    const body = doc.getBody();

    body.appendParagraph('QAHub Status Report').setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph(`Generated ${Utils.formatDateLabel(now)}`);

    body.appendParagraph('Overall Quality Status').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(`${data.qualityStatus} — ${data.qualityReason}`);
    body.appendParagraph(`Comment: ${comment || 'None'}`);

    body.appendParagraph(`Throughput (last ${CONFIG.REPORT.LOOKBACK_DAYS} days)`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(`${data.createdCount} issue(s) created, ${data.closedCount} issue(s) closed.`);

    body.appendParagraph(`Critical+ Issues Found & Fixed (last ${CONFIG.REPORT.LOOKBACK_DAYS} days)`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);

    appendIssueTable(
      body,
      data.foundAndFixed,
      ['Bug ID', 'Title', 'Severity', 'Created', 'Closed'],
      row => [String(row.bugId), row.title, row.severity, Utils.formatDateLabel(row.createdDate), Utils.formatDateLabel(row.closedDate)]
    );

    body.appendParagraph('Critical+ Issues Still Open').setHeading(DocumentApp.ParagraphHeading.HEADING1);

    appendIssueTable(
      body,
      data.criticalOpen,
      ['Bug ID', 'Title', 'Status', 'Severity', 'Assigned To', 'Notes'],
      row => [String(row.bugId), row.title, row.status, row.severity, row.assignedTo || '(unassigned)', row.notes || '']
    );

    body.appendParagraph('Backlog').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(`${data.notInProgressOrClosedCount} open issue(s) are not In Progress or Closed.`);

    doc.saveAndClose();

    moveToSpreadsheetFolder(DriveApp.getFileById(doc.getId()));

    return doc;

  }

  /**
   * Appends a header+data table, or a "None." paragraph if there are no rows.
   */
  function appendIssueTable(body, rows, headers, rowMapper) {

    if (rows.length === 0) {
      body.appendParagraph('None.');
      return;
    }

    body.appendTable([headers].concat(rows.map(rowMapper)));

  }

  /**
   * Moves a Drive file into the same folder as the spreadsheet (Docs are
   * created in Drive's root by default). No-op if the spreadsheet has no
   * parent folder (e.g. it's already at Drive's root).
   */
  function moveToSpreadsheetFolder(file) {

    const spreadsheetFile = DriveApp.getFileById(Utils.spreadsheet().getId());

    const parents = spreadsheetFile.getParents();

    if (!parents.hasNext()) {
      return;
    }

    const folder = parents.next();

    folder.addFile(file);

    DriveApp.getRootFolder().removeFile(file);

  }

  return {
    generate,
    buildReportData,
    suggestQualityStatus,
    isCriticalOrHigher
  };

})();
