/*
 * QAHub
 * Version: 1.0.0
 *
 * File:
 * SearchService.gs
 */

const SearchService = (() => {

  // Columns checked for a keyword match.
  const SEARCH_COLUMNS = [
    CONFIG.COLUMNS.BUG_ID,
    CONFIG.COLUMNS.TITLE,
    CONFIG.COLUMNS.DESCRIPTION,
    CONFIG.COLUMNS.KEYWORDS
  ];

  /**
   * Prompts for a keyword, then jumps to the first matching Issues row.
   */
  function search() {

    const ui = SpreadsheetApp.getUi();

    const response = ui.prompt(
      'Search Issues',
      'Enter a keyword to search Bug ID, Title, Description, and Keywords:',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    const term = response.getResponseText().trim();

    if (!term) {
      return;
    }

    const matches = findMatches(term);

    if (matches.length === 0) {
      ui.alert(`No matches found for "${term}".`);
      return;
    }

    const sheet = Utils.getIssuesSheet();
    const firstMatchRow = matches[0];

    sheet.setActiveRange(sheet.getRange(firstMatchRow, 1));

    ui.alert(
      matches.length === 1
        ? `Found 1 match. Jumped to row ${firstMatchRow}.`
        : `Found ${matches.length} matches. Jumped to the first at row ${firstMatchRow}.`
    );

  }

  /**
   * Returns the sheet row numbers of every Issues row whose Bug ID, Title,
   * Description, or Keywords contain the given term (case-insensitive).
   *
   * @param {string} term
   * @returns {number[]}
   */
  function findMatches(term) {

    const rows = Utils.readDataRows(Utils.getIssuesSheet());

    const needle = term.toLowerCase();

    const matches = [];

    rows.forEach((row, index) => {

      const isMatch = SEARCH_COLUMNS.some(column =>
        String(row[column - 1]).toLowerCase().includes(needle)
      );

      if (isMatch) {
        matches.push(CONFIG.FIRST_DATA_ROW + index);
      }

    });

    return matches;

  }

  return {
    search,
    findMatches
  };

})();
