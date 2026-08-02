/*
 * ValidationService.gs
 */

const ValidationService = (() => {

  function initialize() {

    const sheet = Utils.getIssuesSheet();

    const lastRow = Math.max(
      sheet.getLastRow(),
      CONFIG.FIRST_DATA_ROW
    );

    Object.values(CONFIG.VALIDATION).forEach(validation => {

      applyValidation(
        sheet,
        validation.issueColumn,
        validation.listColumn,
        CONFIG.FIRST_DATA_ROW,
        lastRow
      );

    });

  }

  function applyToRow(sheet, row) {

    Object.values(CONFIG.VALIDATION).forEach(validation => {

      applyValidation(
        sheet,
        validation.issueColumn,
        validation.listColumn,
        row,
        row
      );

    });

  }

  function applyValidation(
    issueSheet,
    issueColumnLetter,
    listColumnLetter,
    startRow,
    endRow
  ) {

    const lists = Utils.getSheet(CONFIG.SHEETS.LISTS);

    const lastListRow = lists.getLastRow();

    const rule = SpreadsheetApp
      .newDataValidation()
      .requireValueInRange(
        lists.getRange(
          `${listColumnLetter}2:${listColumnLetter}${lastListRow}`
        ),
        true
      )
      .setAllowInvalid(false)
      .build();

    const issueColumn =
      Utils.columnLetterToNumber(issueColumnLetter);

    issueSheet
      .getRange(
        startRow,
        issueColumn,
        endRow - startRow + 1,
        1
      )
      .setDataValidation(rule);

  }

  return {

    initialize,

    applyToRow

  };

})();