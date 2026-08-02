/*
 * QAHub
 * Version: 1.0.0
 *
 * File:
 * IssueService.gs
 */

const IssueService = (() => {

  /**
   * Creates a new issue.
   */
  function createIssue() {

    const sheet = Utils.getIssuesSheet();

    // Insert a new row directly beneath the header.
    sheet.insertRowBefore(CONFIG.FIRST_DATA_ROW);

    const row = CONFIG.FIRST_DATA_ROW;

    const now = Utils.now();

    const bugId = Utils.nextBugId();

    const creator = Utils.currentUser();

    const values = [];

    values[CONFIG.COLUMNS.BUG_ID - 1] = bugId;
    values[CONFIG.COLUMNS.STATUS - 1] = CONFIG.STATUS.DEFAULT;
    values[CONFIG.COLUMNS.PRIORITY - 1] = '';
    values[CONFIG.COLUMNS.SEVERITY - 1] = '';
    values[CONFIG.COLUMNS.TITLE - 1] = '';
    values[CONFIG.COLUMNS.DESCRIPTION - 1] = CONFIG.TEMPLATES.DESCRIPTION;
    values[CONFIG.COLUMNS.STEPS - 1] = CONFIG.TEMPLATES.STEPS;
    values[CONFIG.COLUMNS.REPRODUCIBILITY - 1] = '';
    values[CONFIG.COLUMNS.CREATOR - 1] = creator;
    values[CONFIG.COLUMNS.ASSIGNED_TO - 1] = '';
    values[CONFIG.COLUMNS.CREATED_DATE - 1] = now;
    values[CONFIG.COLUMNS.LAST_UPDATED - 1] = now;
    values[CONFIG.COLUMNS.FIXED_IN_BUILD - 1] = '';
    values[CONFIG.COLUMNS.KEYWORDS - 1] = '';
    values[CONFIG.COLUMNS.SCREENSHOT - 1] = '';
    values[CONFIG.COLUMNS.NOTES - 1] = '';

    sheet
      .getRange(row, 1, 1, values.length)
      .setValues([values]);

    Utils.formatIssueRow(sheet, row);

    Utils.focusTitle(sheet, row);

  }

  return {
    createIssue
  };

})();