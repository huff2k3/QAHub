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
    .addSeparator()
    .addItem(CONFIG.MENU.DASHBOARD, 'refreshDashboard')
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
 * Prompts for a keyword and jumps to the first matching issue.
 */
function showSearch() {

  try {

    SearchService.search();

  } catch (error) {

    SpreadsheetApp.getUi().alert(
      'Unable to search issues.\n\n' + error.message
    );

    console.error(error);

  }

}


/**
 * Moves every Closed issue onto the Archive sheet.
 */
function archiveClosed() {

  try {

    const count = ArchiveService.archiveClosed();

    SpreadsheetApp.getUi().alert(
      count === 0
        ? 'No closed issues to archive.'
        : `Archived ${count} closed issue${count === 1 ? '' : 's'}.`
    );

  } catch (error) {

    SpreadsheetApp.getUi().alert(
      'Unable to archive closed issues.\n\n' + error.message
    );

    console.error(error);

  }

}

/**
 * Rebuilds the Dashboard sheet's breakdowns, throughput, and stale-issue list.
 */
function refreshDashboard() {

  try {

    DashboardService.refresh();

  } catch (error) {

    SpreadsheetApp.getUi().alert(
      'Unable to refresh the dashboard.\n\n' + error.message
    );

    console.error(error);

  }

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