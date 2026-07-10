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

    const sheet = getIssuesSheet();

    if (!sheet) {
      throw new Error('Issues sheet not found.');
    }

    // Insert a new row directly beneath the header.
    sheet.insertRowBefore(CONFIG.FIRST_DATA_ROW);

    const row = CONFIG.FIRST_DATA_ROW;

    const now = new Date();

    const bugId = getNextBugId();

    const creator = getCurrentUser();

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

    formatIssueRow(sheet, row);

    // Put the cursor in the Title field.
    sheet.setActiveRange(
      sheet.getRange(
        row,
        CONFIG.COLUMNS.TITLE
      )
    );

  }

  /**
   * Applies formatting to a newly created issue.
   */
  function formatIssueRow(sheet, row) {

    sheet
      .getRange(row, CONFIG.COLUMNS.CREATED_DATE)
      .setNumberFormat("M/d/yyyy h:mm AM/PM");

    sheet
      .getRange(row, CONFIG.COLUMNS.LAST_UPDATED)
      .setNumberFormat("M/d/yyyy h:mm AM/PM");

    sheet
      .getRange(row, 1, 1, 16)
      .setVerticalAlignment("top");

    sheet
      .getRange(row, 1, 1, 16)
      .setWrap(true);

  }

  /**
   * Attempts to determine the current user.
   */
  function getCurrentUser() {

    try {

      const email = Session.getActiveUser().getEmail();

      if (email) {
        return email;
      }

    } catch (e) {
      // Ignore.
    }

    return '';

  }

  return {
    createIssue
  };

})();