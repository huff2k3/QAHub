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
    SETTINGS: 'Settings',
    DASHBOARD: 'Dashboard'
  },

  SCRIPT_PROPERTIES: {
    NEXT_BUG_ID: 'NEXT_BUG_ID'
  },

  MENU: {
    TITLE: 'QA Tools',
    NEW_ISSUE: 'Add New Issue',
    SEARCH: 'Search',
    ARCHIVE: 'Archive Closed',
    DASHBOARD: 'Refresh Dashboard'
  },

  DASHBOARD: {
    // An open (non-Closed) issue whose Last Updated is at least this many
    // days old shows up in the dashboard's stale-issues section.
    STALE_DAYS: 14
  },

  STATUS: {
    DEFAULT: 'New',
    CLOSED: 'Closed'
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
    NOTES: 16,
    CLOSED_DATE: 17

  }),

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

  HEADERS: Object.freeze({

    BUG_ID: 'Bug ID',
    STATUS: 'Status',
    PRIORITY: 'Priority',
    SEVERITY: 'Severity',
    TITLE: 'Title',
    DESCRIPTION: 'Description',
    STEPS: 'Steps to Reproduce',
    REPRODUCIBILITY: 'Reproducibility',
    CREATOR: 'Creator',
    ASSIGNED_TO: 'Assigned To',
    CREATED_DATE: 'Created Date',
    LAST_UPDATED: 'Last Updated',
    FIXED_IN_BUILD: 'Fixed in Build',
    KEYWORDS: 'Keywords',
    SCREENSHOT: 'Screenshot',
    NOTES: 'Notes',
    CLOSED_DATE: 'Closed Date'

  }),

  // Reference values seeded onto the Lists sheet and used by ValidationService
  // for Issues-sheet dropdowns. `column` must match the corresponding
  // `listColumn` in CONFIG.VALIDATION above.
  LISTS: Object.freeze({

    STATUS: {
      column: 'A',
      header: 'Status',
      values: Object.freeze([
        'New',
        'In Progress',
        'Need More Info',
        'Ready for Testing',
        'Closed'
      ])
    },

    SEVERITY: {
      column: 'B',
      header: 'Severity',
      values: Object.freeze([
        'Blocking',
        'Critical',
        'High',
        'Medium',
        'Low'
      ])
    },

    PRIORITY: {
      column: 'C',
      header: 'Priority',
      values: Object.freeze([
        'P1',
        'P2',
        'P3',
        'P4'
      ])
    },

    REPRODUCIBILITY: {
      column: 'D',
      header: 'Reproducibility',
      values: Object.freeze([
        'Once',
        '25%',
        '50%',
        '75%',
        'Always'
      ])
    }

  }),

  HEADER_ROW: 1,

  FIRST_DATA_ROW: 2

});
