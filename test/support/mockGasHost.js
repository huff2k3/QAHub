/**
 * Minimal fakes for the Apps Script host services QAHub touches, sufficient
 * to unit test the business logic in Node/Jest without a real Spreadsheet.
 */

function createRange() {

  const range = {};

  [
    'setValue',
    'setValues',
    'setNumberFormat',
    'setWrap',
    'setVerticalAlignment',
    'setHorizontalAlignment',
    'setFontWeight',
    'setFontSize',
    'setDataValidation',
    'createFilter'
  ].forEach(method => {
    range[method] = jest.fn(() => range);
  });

  range.getValues = jest.fn(() => []);

  return range;

}

function createChartBuilder() {

  const builder = {};

  builder.setChartType = jest.fn(() => builder);
  builder.addRange = jest.fn(() => builder);
  builder.setPosition = jest.fn(() => builder);
  builder.setOption = jest.fn(() => builder);
  builder.build = jest.fn(() => ({ __chart: true }));

  return builder;

}

function createSheet(name, overrides = {}) {

  const sharedRange = createRange();

  return {
    name,
    range: sharedRange,
    getName: jest.fn(() => name),
    getRange: jest.fn(() => sharedRange),
    insertRowBefore: jest.fn(),
    deleteRow: jest.fn(),
    clear: jest.fn(),
    setActiveRange: jest.fn(),
    setActiveSelection: jest.fn(),
    setFrozenRows: jest.fn(),
    setColumnWidth: jest.fn(),
    getFilter: jest.fn(() => null),
    getLastRow: jest.fn(() => 1),
    getMaxRows: jest.fn(() => 1000),
    getLastColumn: jest.fn(() => 16),
    getCharts: jest.fn(() => []),
    removeChart: jest.fn(),
    newChart: jest.fn(() => createChartBuilder()),
    insertChart: jest.fn(),
    ...overrides
  };

}

function createDataValidationBuilder() {

  const builder = {};

  builder.requireValueInRange = jest.fn(() => builder);
  builder.setAllowInvalid = jest.fn(() => builder);
  builder.build = jest.fn(() => ({ __rule: true }));

  return builder;

}

function createCharts() {
  return {
    ChartType: { COLUMN: 'COLUMN', PIE: 'PIE', LINE: 'LINE' }
  };
}

function createUi(promptResponse = { button: 'CANCEL', text: '' }) {

  return {
    createMenu: jest.fn(() => ({
      addItem: jest.fn(function () { return this; }),
      addSeparator: jest.fn(function () { return this; }),
      addToUi: jest.fn()
    })),
    alert: jest.fn(),
    prompt: jest.fn(() => ({
      getSelectedButton: jest.fn(() => promptResponse.button),
      getResponseText: jest.fn(() => promptResponse.text)
    })),
    Button: { OK: 'OK', CANCEL: 'CANCEL' },
    ButtonSet: { OK_CANCEL: 'OK_CANCEL' }
  };

}

function createSpreadsheetApp(sheetsByName, options = {}) {

  const spreadsheet = {
    getId: jest.fn(() => options.spreadsheetId || 'spreadsheet-id'),
    getSheetByName: jest.fn(sheetName => sheetsByName[sheetName] || null),
    insertSheet: jest.fn(sheetName => {
      const sheet = createSheet(sheetName);
      sheetsByName[sheetName] = sheet;
      return sheet;
    })
  };

  const ui = createUi(options.promptResponse);

  return {
    getActiveSpreadsheet: jest.fn(() => spreadsheet),
    newDataValidation: jest.fn(() => createDataValidationBuilder()),
    getUi: jest.fn(() => ui)
  };

}

/**
 * Fakes DocumentApp: create() returns a Doc whose Body records every
 * appendParagraph/appendTable call, enough to assert on report structure
 * without a real Google Doc.
 */
function createDocumentApp() {

  const created = [];

  const documentApp = {
    create: jest.fn(name => {

      const paragraphs = [];

      const body = {
        appendParagraph: jest.fn(text => {
          const paragraph = { text, heading: null };
          paragraph.setHeading = jest.fn(heading => { paragraph.heading = heading; return paragraph; });
          paragraphs.push(paragraph);
          return paragraph;
        }),
        appendTable: jest.fn(cells => { body.tables = body.tables || []; body.tables.push(cells); return {}; }),
        paragraphs
      };

      const doc = {
        name,
        getId: jest.fn(() => 'doc-id'),
        getUrl: jest.fn(() => `https://docs.google.com/document/d/doc-id/edit`),
        getBody: jest.fn(() => body),
        saveAndClose: jest.fn(),
        body
      };

      created.push(doc);

      return doc;

    }),
    ParagraphHeading: { TITLE: 'TITLE', HEADING1: 'HEADING1', HEADING2: 'HEADING2' },
    created
  };

  return documentApp;

}

/**
 * Fakes DriveApp: getFileById(id) looks up from a caller-provided map, so a
 * test can register both the spreadsheet's Drive file (with a parent
 * folder, or not) and the report Doc's Drive file.
 */
function createDriveApp(filesById = {}) {

  const rootFolder = { removeFile: jest.fn() };

  return {
    getFileById: jest.fn(id => filesById[id] || { getParents: jest.fn(() => ({ hasNext: () => false })) }),
    getRootFolder: jest.fn(() => rootFolder),
    rootFolder
  };

}

/**
 * A fake Drive file. Pass a folder (from createDriveFolder()) if it should
 * report having a parent; omit for a file at Drive's root.
 */
function createDriveFile(parentFolder) {

  return {
    getParents: jest.fn(() => ({
      hasNext: jest.fn(() => !!parentFolder),
      next: jest.fn(() => parentFolder)
    }))
  };

}

function createDriveFolder() {
  return { addFile: jest.fn() };
}

function createPropertiesService(initial = {}) {

  const store = { ...initial };

  return {
    getScriptProperties: jest.fn(() => ({
      getProperty: jest.fn(key => (key in store ? store[key] : null)),
      setProperty: jest.fn((key, value) => { store[key] = value; })
    }))
  };

}

function createLockService() {

  return {
    getScriptLock: jest.fn(() => ({
      waitLock: jest.fn(),
      releaseLock: jest.fn()
    }))
  };

}

function createSession(email = 'tester@example.com') {

  return {
    getActiveUser: jest.fn(() => ({
      getEmail: jest.fn(() => email)
    }))
  };

}

module.exports = {
  createRange,
  createSheet,
  createUi,
  createSpreadsheetApp,
  createPropertiesService,
  createLockService,
  createSession,
  createCharts,
  createDocumentApp,
  createDriveApp,
  createDriveFile,
  createDriveFolder
};
