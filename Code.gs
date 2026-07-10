/*
 * QAHub
 * Version: 1.0.0
 *
 * File:
 * Code.gs
 *
 * Application entry points and custom menu.
 */

/**
 * Runs whenever the spreadsheet is opened.
 */
function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu(CONFIG.MENU.TITLE)
    .addItem(CONFIG.MENU.NEW_ISSUE, 'addNewIssue')
    .addSeparator()
    .addItem('Initialize Workbook', 'initializeWorkbook')
    .addSeparator()
    .addItem(CONFIG.MENU.SEARCH, 'showSearch')
    .addItem(CONFIG.MENU.ARCHIVE, 'archiveClosed')
    .addToUi();

}


/**
 * Creates a new issue.
 */
function addNewIssue() {

  try {

    IssueService.createIssue();

  } catch (error) {

    SpreadsheetApp.getUi().alert(
      'Unable to create a new issue.\n\n' + error.message
    );

    console.error(error);

  }

}


/**
 * Opens the Search dialog.
 *
 * Placeholder until SearchService is implemented.
 */
function showSearch() {

  SpreadsheetApp.getUi().alert(
    'Search is coming in Version 1.1.'
  );

}


/**
 * Archives closed issues.
 *
 * Placeholder until ArchiveService is implemented.
 */
function archiveClosed() {

  SpreadsheetApp.getUi().alert(
    'Archive is coming in Version 1.1.'
  );

}

function initializeWorkbook() {

  try {

    SetupService.initializeWorkbook();

    SpreadsheetApp.getUi().alert(
      "QAHub initialized successfully."
    );

  } catch (error) {

    SpreadsheetApp.getUi().alert(error.message);

  }

}