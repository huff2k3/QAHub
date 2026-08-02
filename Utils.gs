/*
 * QAHub
 * Version: 1.0.0
 *
 * File:
 * Utils.gs
 *
 * Shared utility functions.
 */

const Utils = (() => {

  /**
   * Returns the active spreadsheet.
   *
   * @returns {Spreadsheet}
   */
  function spreadsheet() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  /**
   * Returns a sheet by name.
   *
   * @param {string} sheetName
   * @returns {Sheet}
   */
  function getSheet(sheetName) {

    const sheet = spreadsheet().getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" was not found.`);
    }

    return sheet;

  }

  /**
   * Returns the Issues sheet.
   *
   * @returns {Sheet}
   */
  function getIssuesSheet() {
    return getSheet(CONFIG.SHEETS.ISSUES);
  }

  /**
   * Returns the Archive sheet.
   *
   * @returns {Sheet}
   */
  function getArchiveSheet() {
    return getSheet(CONFIG.SHEETS.ARCHIVE);
  }

  /**
   * Returns script properties.
   *
   * @returns {Properties}
   */
  function scriptProperties() {
    return PropertiesService.getScriptProperties();
  }

  /**
   * Returns the active user's email when available.
   *
   * May be blank depending on Google Workspace settings.
   *
   * @returns {string}
   */
  function currentUser() {

    try {

      const email = Session.getActiveUser().getEmail();

      return email || '';

    } catch (e) {

      return '';

    }

  }

  /**
   * Returns the current timestamp.
   *
   * @returns {Date}
   */
  function now() {
    return new Date();
  }

  /**
   * Applies standard formatting to an issue row.
   *
   * @param {Sheet} sheet
   * @param {number} row
   */
  function formatIssueRow(sheet, row) {

    const totalColumns = Object.keys(CONFIG.COLUMNS).length;

    sheet
      .getRange(row, 1, 1, totalColumns)
      .setWrap(true)
      .setVerticalAlignment("top");

    sheet
      .getRange(row, CONFIG.COLUMNS.CREATED_DATE)
      .setNumberFormat("M/d/yyyy h:mm AM/PM");

    sheet
      .getRange(row, CONFIG.COLUMNS.LAST_UPDATED)
      .setNumberFormat("M/d/yyyy h:mm AM/PM");

  }

  /**
   * Moves the cursor to the Title column.
   *
   * @param {Sheet} sheet
   * @param {number} row
   */
  function focusTitle(sheet, row) {

    sheet.setActiveSelection(
      sheet.getRange(
        row,
        CONFIG.COLUMNS.TITLE
      )
    );

  }

  /**
   * Returns the next Bug ID.
   *
   * Uses LockService to avoid duplicate IDs.
   *
   * @returns {number}
   */
  function nextBugId() {

    const lock = LockService.getScriptLock();

    lock.waitLock(30000);

    try {

      const props = scriptProperties();

      let nextId = Number(
        props.getProperty(CONFIG.SCRIPT_PROPERTIES.NEXT_BUG_ID)
      );

      if (!nextId || isNaN(nextId)) {
        nextId = 1;
      }

      props.setProperty(
        CONFIG.SCRIPT_PROPERTIES.NEXT_BUG_ID,
        String(nextId + 1)
      );

      return nextId;

    } finally {

      lock.releaseLock();

    }

  }

  /**
   * Updates the Last Updated timestamp.
   *
   * @param {Sheet} sheet
   * @param {number} row
   */
  function touch(sheet, row) {

    sheet
      .getRange(
        row,
        CONFIG.COLUMNS.LAST_UPDATED
      )
      .setValue(now());

  }

  /**
   * Converts a spreadsheet column letter (e.g. "A", "H") to its 1-based
   * column number.
   *
   * @param {string} letter
   * @returns {number}
   */
  function columnLetterToNumber(letter) {

    let column = 0;

    for (let i = 0; i < letter.length; i++) {

      column *= 26;
      column += letter.charCodeAt(i) - 64;

    }

    return column;

  }

  return {

    spreadsheet,

    getSheet,

    getIssuesSheet,

    getArchiveSheet,

    scriptProperties,

    currentUser,

    now,

    nextBugId,

    formatIssueRow,

    focusTitle,

    touch,

    columnLetterToNumber

  };

})();