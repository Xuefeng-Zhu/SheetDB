const fs = require('fs');
const vm = require('vm');

function loadUtils() {
  const code = fs.readFileSync('src/Utils.gs', 'utf8');
  const sandbox = {
    Utilities: {
      formatDate: () => '2026-01-02 03:04:05',
    },
    Session: {
      getScriptTimeZone: () => 'UTC',
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(code, sandbox);
  return sandbox;
}

function assertEqual(actual, expected, message) {
  if (actual !== expected) {
    throw new Error(`${message}\nExpected: ${expected}\nActual:   ${actual}`);
  }
}

(function run() {
  const u = loadUtils();

  assertEqual(u.isCancel('cancel'), true, 'isCancel should recognize cancel');
  assertEqual(u.isCancel('ok'), false, 'isCancel should reject non-cancel input');

  assertEqual(u.quoteSqlValue(null), 'NULL', 'quoteSqlValue should convert null to NULL');
  assertEqual(u.quoteSqlValue(''), 'NULL', 'quoteSqlValue should convert empty string to NULL');
  assertEqual(u.quoteSqlValue(123), '123', 'quoteSqlValue should preserve numbers');
  assertEqual(u.quoteSqlValue("O'Reilly"), "'O''Reilly'", 'quoteSqlValue should escape single quotes');

  console.log('All tests passed.');
})();
