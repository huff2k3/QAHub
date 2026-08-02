/*
 * QAHub
 * Version: 1.0.0
 *
 * File:
 * TriggerService.gs
 */

function onEdit(e) {

  if (!e) {
    return;
  }

  const sheet = e.range.getSheet();

  if (sheet.getName() !== CONFIG.SHEETS.ISSUES) {
    return;
  }

  const row = e.range.getRow();

  if (row < CONFIG.FIRST_DATA_ROW) {
    return;
  }

  const column = e.range.getColumn();

  // Ignore edits to Last Updated itself.
  if (column === CONFIG.COLUMNS.LAST_UPDATED) {
    return;
  }

  if (column === CONFIG.COLUMNS.STATUS) {
    updateClosedDate(sheet, row, e.value);
  }

  updateLastUpdated(sheet, row);

}


/**
 * Updates the Last Updated timestamp.
 */
function updateLastUpdated(sheet, row) {

  sheet
    .getRange(
      row,
      CONFIG.COLUMNS.LAST_UPDATED
    )
    .setValue(new Date());

}


/**
 * Stamps Closed Date when Status is set to Closed, clears it otherwise
 * (e.g. an issue gets reopened) so it only ever reflects the current
 * closure, not a stale one.
 */
function updateClosedDate(sheet, row, newStatus) {

  sheet
    .getRange(
      row,
      CONFIG.COLUMNS.CLOSED_DATE
    )
    .setValue(newStatus === CONFIG.STATUS.CLOSED ? new Date() : '');

}