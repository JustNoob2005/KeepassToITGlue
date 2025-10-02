// Configuration and Script Properties helpers

/** Auto-generated restructure. */


/**
 * Returns a script property by key.
 * @param {string} key
 * @returns {string|null}
 */
function CONFIG_get(key) {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty(key);
}

/**
 * Returns the active spreadsheet.
 * Use this instead of a global `ss` variable.
 */
function SS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Returns a script-scoped lock.
 * Use this instead of a global `lock` variable.
 */
function LOCK() {
  return LockService.getScriptLock();
}
