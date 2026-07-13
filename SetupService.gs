/*
 * QAHub
 * Version: 1.0.0
 *
 * File:
 * SetupService.gs
 */

const SetupService = (() => {

  /**
   * Initializes the workbook.
   */
  function initializeWorkbook() {

    const sheet = Utils.getIssuesSheet();

    freezeHeader(sheet);
    enableFilters(sheet);
    formatHeader(sheet);
    autoResize(sheet);

    ValidationService.initialize();

  }

  /**
   * Freeze header row.
   */
  function freezeHeader(sheet) {
    sheet.setFrozenRows(1);
  }

  /**
   * Enable filters.
   */
  function enableFilters(sheet) {

    if (!sheet.getFilter()) {
      sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getLastColumn())
           .createFilter();
    }

  }

  /**
   * Format header.
   */
  function formatHeader(sheet) {

    const header = sheet.getRange(
      1,
      1,
      1,
      Object.keys(CONFIG.COLUMNS).length
    );

    header
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);

  }

  /**
   * Set default column widths.
   */
  function autoResize(sheet) {

    const widths = {
      1: 90,
      2: 140,
      3: 90,
      4: 100,
      5: 350,
      6: 300,
      7: 350,
      8: 120,
      9: 160,
      10: 160,
      11: 150,
      12: 150,
      13: 120,
      14: 200,
      15: 250,
      16: 300
    };

    Object.keys(widths).forEach(column => {
      sheet.setColumnWidth(Number(column), widths[column]);
    });

  }

  return {
    initializeWorkbook
  };

})();