const { loadGasProject } = require('./support/gasEnv');
const { createSpreadsheetApp, createPropertiesService, createLockService, createSession } = require('./support/mockGasHost');

describe('CONFIG', () => {

  function load() {
    return loadGasProject({
      SpreadsheetApp: createSpreadsheetApp({}),
      PropertiesService: createPropertiesService(),
      LockService: createLockService(),
      Session: createSession()
    });
  }

  test('parses without throwing', () => {
    expect(load).not.toThrow();
  });

  test('is frozen', () => {
    const { CONFIG } = load();
    expect(Object.isFrozen(CONFIG)).toBe(true);
  });

  test('exposes a VALIDATION map with issue/list column pairs', () => {
    const { CONFIG } = load();

    expect(CONFIG.VALIDATION).toBeDefined();
    expect(CONFIG.VALIDATION.STATUS).toEqual({ issueColumn: 'B', listColumn: 'A' });
    expect(CONFIG.VALIDATION.SEVERITY).toEqual({ issueColumn: 'D', listColumn: 'B' });
    expect(CONFIG.VALIDATION.PRIORITY).toEqual({ issueColumn: 'C', listColumn: 'C' });
    expect(CONFIG.VALIDATION.REPRODUCIBILITY).toEqual({ issueColumn: 'H', listColumn: 'D' });
  });

  test('COLUMNS covers all 16 fields with unique 1-based indexes', () => {
    const { CONFIG } = load();

    const indexes = Object.values(CONFIG.COLUMNS);

    expect(indexes).toHaveLength(16);
    expect(new Set(indexes).size).toBe(16);
    expect(Math.min(...indexes)).toBe(1);
    expect(Math.max(...indexes)).toBe(16);
  });

});
