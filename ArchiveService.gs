/*
 * QAHub
 * Version: 1.0.0
 *
 * File:
 * ArchiveService.gs
 */

const ArchiveService = (() => {

  /**
   * Moves every Closed issue from the Issues sheet onto the Archive sheet.
   *
   * @returns {number} How many issues were archived.
   */
  function archiveClosed() {

    const issuesSheet = Utils.getIssuesSheet();
    const archiveSheet = Utils.getArchiveSheet();

    const lastRow = issuesSheet.getLastRow();

    if (lastRow < CONFIG.FIRST_DATA_ROW) {
      return 0;
    }

    const totalColumns = Object.keys(CONFIG.COLUMNS).length;

    const rows = issuesSheet
      .getRange(CONFIG.FIRST_DATA_ROW, 1, lastRow - CONFIG.FIRST_DATA_ROW + 1, totalColumns)
      .getValues();

    let archivedCount = 0;

    // Walk bottom-up so deleting a row doesn't shift the sheet row number of
    // any row still left to check.
    for (let i = rows.length - 1; i >= 0; i--) {

      const row = rows[i];

      if (row[CONFIG.COLUMNS.STATUS - 1] !== CONFIG.STATUS.CLOSED) {
        continue;
      }

      appendRow(archiveSheet, row);

      issuesSheet.deleteRow(CONFIG.FIRST_DATA_ROW + i);

      archivedCount++;

    }

    return archivedCount;

  }

  /**
   * Appends a row of values to the first empty row of a sheet.
   */
  function appendRow(sheet, values) {

    const nextRow = sheet.getLastRow() + 1;

    sheet
      .getRange(nextRow, 1, 1, values.length)
      .setValues([values]);

  }

  return {
    archiveClosed
  };

})();
