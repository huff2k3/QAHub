/*
 * QAHub
 * Version: 1.0.0
 *
 * File:
 * SetupService.gs
 */

const SetupService = (() => {

  /**
   * Builds a fresh workbook: creates the Issues/Archive/Lists/Settings
   * sheets if missing, writes headers, seeds the Lists reference values,
   * formats the Issues and Archive sheets, and applies data validation.
   *
   * Safe to re-run: re-running resets headers and Lists values back to the
   * configured defaults (overwriting any manual edits to those specifically),
   * but never touches issue data rows.
   */
  function initializeWorkbook() {

    const issuesSheet = Utils.ensureSheet(CONFIG.SHEETS.ISSUES);
    const archiveSheet = Utils.ensureSheet(CONFIG.SHEETS.ARCHIVE);
    const listsSheet = Utils.ensureSheet(CONFIG.SHEETS.LISTS);

    Utils.ensureSheet(CONFIG.SHEETS.SETTINGS);
    Utils.ensureSheet(CONFIG.SHEETS.DASHBOARD);

    writeHeaders(issuesSheet);
    writeHeaders(archiveSheet);
    seedLists(listsSheet);

    [issuesSheet, archiveSheet].forEach(sheet => {

      freezeHeader(sheet);
      enableFilters(sheet);
      formatHeader(sheet);
      autoResize(sheet);

    });

    ValidationService.initialize();

  }

  /**
   * Writes the Issues column headers, in CONFIG.COLUMNS order.
   */
  function writeHeaders(sheet) {

    const headers = Object.keys(CONFIG.COLUMNS)
      .sort((a, b) => CONFIG.COLUMNS[a] - CONFIG.COLUMNS[b])
      .map(key => CONFIG.HEADERS[key]);

    sheet
      .getRange(CONFIG.HEADER_ROW, 1, 1, headers.length)
      .setValues([headers]);

  }

  /**
   * Writes each configured reference list's header and values onto the
   * Lists sheet, in the column CONFIG.VALIDATION expects it in.
   */
  function seedLists(sheet) {

    Object.values(CONFIG.LISTS).forEach(list => {

      const column = Utils.columnLetterToNumber(list.column);

      sheet
        .getRange(CONFIG.HEADER_ROW, column)
        .setValue(list.header);

      sheet
        .getRange(CONFIG.FIRST_DATA_ROW, column, list.values.length, 1)
        .setValues(list.values.map(value => [value]));

    });

  }

  /**
   * Freeze header row.
   */
  function freezeHeader(sheet) {
    sheet.setFrozenRows(1);
  }

  /**
   * (Re)creates the sheet's filter over the full current column range.
   *
   * Removes any existing filter first — a filter's range doesn't grow on
   * its own when a new column is added to CONFIG.COLUMNS, so on re-run
   * this is what actually extends the filter to cover it. Recreating
   * clears any filter criteria the user had set, same tradeoff as
   * initializeWorkbook resetting headers/Lists on re-run.
   */
  function enableFilters(sheet) {

    const existingFilter = sheet.getFilter();

    if (existingFilter) {
      existingFilter.remove();
    }

    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getLastColumn())
         .createFilter();

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
      16: 300,
      17: 150
    };

    Object.keys(widths).forEach(column => {
      sheet.setColumnWidth(Number(column), widths[column]);
    });

  }

  return {
    initializeWorkbook
  };

})();
