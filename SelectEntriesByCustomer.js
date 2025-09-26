/**
 * Uses the FIRST data row in "Klant import" to determine the customer
 * (from the path "#klanten/{customer}[...]"),
 * finds that customer's ID in "Klanten overzicht" (group "Klanten"),
 * and returns all "Klant import" rows for that customer.
 *
 * @returns {{ customerName: string, customerId: string|null, matches: Object[] }}
 */
function selectFirstCustomerEntries() {
  const ss = SpreadsheetApp.getActive();
  const SH_IMPORT = 'Klant import';
  const SH_OVERVIEW = 'Klanten overzicht';
  const TARGET_GROUP = 'Klanten';

  // --- Read Klant import ---
  const shImport = ss.getSheetByName(SH_IMPORT);
  if (!shImport) throw new Error(`Sheet "${SH_IMPORT}" not found.`);
  const lastRow = shImport.getLastRow();
  const lastCol = shImport.getLastColumn();
  if (lastRow < 2) return { customerName: '', customerId: null, matches: [] };

  // A..F = entry, user, pass, url, notes, path
  const COLS = { entry: 1, user: 2, pass: 3, url: 4, notes: 5, path: 6 };
  const data = shImport.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Determine target customer from the FIRST data row's Path
  const firstPath = String((data[0] && data[0][COLS.path - 1]) || '').trim();
  if (!firstPath) throw new Error('First data row has an empty Path in "Klant import".');

  const customerName = extractCustomerFromPath_(firstPath);
  if (!customerName) {
    throw new Error(`Could not extract customer from first Path: "${firstPath}". Expected "#klanten/{customer}[...]".`);
  }

  // --- Find the customer ID (index in col B) in Klanten overzicht for group "Klanten" ---

  const customerId = findCustomerIdFromOverview_(ss, SH_OVERVIEW, TARGET_GROUP, customerName);

  // --- Collect all rows in Klant import for this customer ---
  const matches = [];
  for (const row of data) {
    const path = String(row[COLS.path - 1] || '').trim();
    const customer = extractCustomerFromPath_(path);
    if (customer && equalsLoose_(customer, customerName)) {
      matches.push({
        entryName: row[COLS.entry - 1],
        username: row[COLS.user - 1],
        password: row[COLS.pass - 1],
        url: row[COLS.url - 1],
        notes: row[COLS.notes - 1],
        path
      });
    }
  }

  Logger.log(`Customer: "${customerName}", ID: ${customerId ?? 'not found'}, matches: ${matches.length}`);
  return { customerName, customerId, matches };
}

/**
 * Look up the customer's ID (index in col B) in "Klanten overzicht"
 * restricted to rows where column A == TARGET_GROUP.
 * Returns the ID as a string, or null if not found.
 */
function findCustomerIdFromOverview_(ss, sheetName, targetGroup, customerName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Sheet "${sheetName}" not found.`);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  // Columns: A=group, B=index (ID), C=name
  const rows = sh.getRange(2, 1, lastRow - 1, 3).getValues();

  const matches = [];
  for (const [group, id, name] of rows) {
    if (isBlank_(group) || isBlank_(name)) continue;
    if (!equalsLoose_(group, targetGroup)) continue;

    if (normalizeName(name).includes(normalizeName(customerName))) {
      if (id !== null && id !== undefined && String(id).trim() !== '') {
        matches.push(String(id).trim());
      }
    }
  }
  if (!matches.length) {

  }
  if (!matches.length) return null;

  // If multiple IDs are present for the same name, pick the first but log it.
  if (matches.length > 1) {
    Logger.log(`Warning: multiple IDs found for "${customerName}" in group "${targetGroup}": [${matches.join(', ')}]. Using first.`);
  }
  return matches[0];
}

/** Extracts "{customer}" from "#klanten/{customer}" or "#klanten/{customer}/..." */
function extractCustomerFromPath_(path) {
  const m = /^#klanten\/([^\/#]+)(?:\/.*)?$/i.exec(path);
  return m ? m[1].trim() : null;
}

/** Case-insensitive, trimmed equality */
function equalsLoose_(a, b) {
  return String(a).trim().toLowerCase() === String(b).trim().toLowerCase();
}

function isBlank_(v) {
  return v === null || v === undefined || String(v).trim() === '';
}
