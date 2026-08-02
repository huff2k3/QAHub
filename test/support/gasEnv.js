const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..', '..');

// Load order matters: later files reference consts declared in earlier ones,
// same as how Apps Script concatenates all .gs files into one global scope.
const FILES = [
  'Config.gs',
  'Utils.gs',
  'ValidationService.gs',
  'IssueService.gs',
  'SetupService.gs',
  'ArchiveService.gs',
  'SearchService.gs',
  'TriggerService.gs',
  'Code.gs'
];

const EXPORTED_NAMES = [
  'CONFIG',
  'Utils',
  'IssueService',
  'SetupService',
  'ValidationService',
  'ArchiveService',
  'SearchService',
  'onOpen',
  'addNewIssue',
  'showSearch',
  'archiveClosed',
  'initializeWorkbook',
  'onEdit',
  'updateLastUpdated'
];

/**
 * Executes the project's .gs files inside a sandbox seeded with fake Apps
 * Script host services, then returns the top-level bindings they define.
 *
 * @param {object} hostServices e.g. { SpreadsheetApp, PropertiesService, LockService, Session }
 * @returns {object}
 */
function loadGasProject(hostServices = {}) {

  const sandbox = { console, ...hostServices };

  vm.createContext(sandbox);

  FILES.forEach(file => {

    const source = fs.readFileSync(path.join(ROOT, file), 'utf8');

    vm.runInContext(source, sandbox, { filename: file });

  });

  // `const`/function bindings at the top level of a vm context live in that
  // context's lexical scope, not as sandbox properties — re-assign them onto
  // a `var` so the test process can read them back off the sandbox object.
  vm.runInContext(
    `var __exports = { ${EXPORTED_NAMES.join(', ')} };`,
    sandbox
  );

  return sandbox.__exports;

}

module.exports = { loadGasProject };
