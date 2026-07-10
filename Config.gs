/*
 * QAHub
 * Version: 1.0.0
 *
 * File:
 * Config.gs
 *
 * Central configuration for QAHub.
 */

const CONFIG = Object.freeze({

  APP: {
    NAME: 'QAHub',
    VERSION: '1.0.0'
  },

  SHEETS: {
    ISSUES: 'Issues',
    ARCHIVE: 'Archive',
    LISTS: 'Lists',
    SETTINGS: 'Settings'
  },

  SCRIPT_PROPERTIES: {
    NEXT_BUG_ID: 'NEXT_BUG_ID'
  },

  MENU: {
    TITLE: 'QA Tools',
    NEW_ISSUE: 'Add New Issue',
    SEARCH: 'Search',
    ARCHIVE: 'Archive Closed'
  },

  STATUS: {
    DEFAULT: 'New'
  },

  TEMPLATES: {

    DESCRIPTION:
`Environment:

Observed Behavior:`,

    STEPS:
`1. Launch game.
2.
3.

Actual Result: 
Expected Result: `
  },

  COLUMNS: Object.freeze({

    BUG_ID: 1,
    STATUS: 2,
    PRIORITY: 3,
    SEVERITY: 4,
    TITLE: 5,
    DESCRIPTION: 6,
    STEPS: 7,
    REPRODUCIBILITY: 8,
    CREATOR: 9,
    ASSIGNED_TO: 10,
    CREATED_DATE: 11,
    LAST_UPDATED: 12,
    FIXED_IN_BUILD: 13,
    KEYWORDS: 14,
    SCREENSHOT: 15,
    NOTES: 16

  }),

  HEADER_ROW: 1,

  FIRST_DATA_ROW: 2

});


/**
 * Returns the Issues sheet.
 */
function getIssuesSheet() {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.ISSUES);
}


/**
 * Returns script properties.
 */
function getScriptProperties() {
  return PropertiesService.getScriptProperties();
}


/**
 * Returns the next Bug ID.
 *
 * Uses LockService to ensure multiple users
 * cannot generate duplicate IDs.
 */
function getNextBugId() {

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    const properties = getScriptProperties();

    let nextId = Number(
      properties.getProperty(CONFIG.SCRIPT_PROPERTIES.NEXT_BUG_ID)
    );

    if (!nextId || isNaN(nextId)) {
      nextId = 1;
    }

    properties.setProperty(
      CONFIG.SCRIPT_PROPERTIES.NEXT_BUG_ID,
      String(nextId + 1)
    );

    return nextId;

  } finally {

    lock.releaseLock();

  }

}

VALIDATION: Object.freeze({

  STATUS: {
    issueColumn: 'B',
    listColumn: 'A'
  },

  SEVERITY: {
    issueColumn: 'D',
    listColumn: 'B'
  },

  PRIORITY: {
    issueColumn: 'C',
    listColumn: 'C'
  },

  REPRODUCIBILITY: {
    issueColumn: 'H',
    listColumn: 'D'
  }

}),