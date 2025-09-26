function testReturn() {

}

function selectEntriesByCustomerIndex(chosenIndex) {
  const ss = SpreadsheetApp.getActive();
  const SH_OVERVIEW = 'Klanten overzicht'; // A=group, B=index, C=name
  const SH_IMPORT = 'Klant import';
  const TARGET_GROUP = 'Klanten';

  const shOverview = ss.getSheetByName(SH_OVERVIEW);
  if (!shOverview) throw new Error(`Sheet "${SH_OVERVIEW}" not found.`);

  const lastRowOv = shOverview.getLastRow();
  if (lastRowOv < 2) throw new Error(`No data in "${SH_OVERVIEW}".`);

  const ovRows = shOverview.getRange(2, 1, lastRowOv - 1, 3).getValues(); // A:C
  const map = new Map();

  for (const [group, idx, name] of ovRows) {
    if (isBlank_(group) || isBlank_(name)) continue;
    if (String(group).trim().toLowerCase() !== TARGET_GROUP.toLowerCase()) continue;

    if (idx === null || idx === undefined) continue;
    const key = String(idx).trim();
    if (key === '') continue;

    const val = String(name).trim();
    if (map.has(key) && map.get(key) !== val) {
      throw new Error(`Duplicate index "${key}" with different names in group "${TARGET_GROUP}".`);
    }
    map.set(key, val);
  }

  const customerName = map.get(String(chosenIndex).trim());
  if (!customerName) {
    throw new Error(`No customer found for index ${chosenIndex} in group "${TARGET_GROUP}".`);
  }

  const shImport = ss.getSheetByName(SH_IMPORT);
  const lastRowIm = shImport.getLastRow();
  const lastColIm = shImport.getLastColumn();
  if (lastRowIm < 2) return [];

  const data = shImport.getRange(2, 1, lastRowIm - 1, lastColIm).getValues();
  const COLS = { entry: 1, user: 2, pass: 3, url: 4, notes: 5, path: 6 };

  const matches = [];
  let pathCustomerForLoose = null;

  // Exact matches
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const path = String(row[COLS.path - 1] || '').trim();
    const pathCustomer = extractCustomerFromPath_(path);
    if (!pathCustomer) continue;

    if (equalsLoose_(pathCustomer, customerName)) {
      const sheetRow = r + 2; // data starts at row 2
      matches.push({
        sheetRow,
        entryName: row[COLS.entry - 1],
        username: row[COLS.user - 1],
        password: row[COLS.pass - 1],
        url: row[COLS.url - 1],
        notes: row[COLS.notes - 1],
        path
      });
    } else if (!pathCustomerForLoose && normalizeName(customerName).includes(normalizeName(pathCustomer))) {
      pathCustomerForLoose = pathCustomer;
    }
  }

  // If no exact matches, ask for loose match
  if (!matches.length && pathCustomerForLoose &&
    normalizeName(customerName).includes(normalizeName(pathCustomerForLoose))) {

    const partMatch = ui.prompt(
      "Deelse match",
      `Klantnaam uit path: "${pathCustomerForLoose}"\nKlantnaam uit overzicht: "${customerName}"\n\nAccepteren?`,
      ui.ButtonSet.OK_CANCEL
    );

    if (partMatch.getSelectedButton() === ui.Button.OK) {
      for (let r = 0; r < data.length; r++) {
        const row = data[r];
        const path = String(row[COLS.path - 1] || '').trim();
        const pc2 = extractCustomerFromPath_(path);
        if (pc2 && normalizeName(customerName).includes(normalizeName(pc2))) {
          const sheetRow = r + 2;
          matches.push({
            sheetRow,
            entryName: row[COLS.entry - 1],
            username: row[COLS.user - 1],
            password: row[COLS.pass - 1],
            url: row[COLS.url - 1],
            notes: row[COLS.notes - 1],
            path
          });
        }
      }
    } else {
      ui.alert("Match geweigerd.");
    }
  }

  // Color the rows we’re going to use
  if (matches.length) {
    matches.forEach(m => ColorSpecificRow(shImport, m.sheetRow, "green")); // or "green"/whatever
  }

  Logger.log(`Group "${TARGET_GROUP}", index ${chosenIndex} → "${customerName}", matches: ${matches.length}`);
  return matches;

  // helpers
  function isBlank_(v) {
    return v === null || v === undefined || String(v).trim() === '';
  }
  function extractCustomerFromPath_(path) {
    const m = /^#klanten\/([^\/#]+)(?:\/.*)?$/i.exec(path);
    return m ? m[1].trim() : null;
  }
  function equalsLoose_(a, b) {
    return String(a).trim().toLowerCase() === String(b).trim().toLowerCase();
  }
  function normalizeName(s) {
    return String(s || "")
      .trim()
      .toLowerCase()
      .replace(/[\/\\]/g, " ")
      .replace(/\s+/g, " ");
  }
}
