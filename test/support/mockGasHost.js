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

function createFilter() {
  return { remove: jest.fn() };
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
  createFilter,
  createUi,
  createSpreadsheetApp,
  createPropertiesService,
  createLockService,
  createSession,
  createCharts
};
