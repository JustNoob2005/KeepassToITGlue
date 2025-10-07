/**
 * Import all ITGlue organization names into the sheet "Klanten overzicht".
 * - Column A: "Klanten"
 * - Column B: index (1..N)
 * - Column C: organization name
 * It clears previous data (from row 2 down) and repopulates.
 */
var ui = SpreadsheetApp.getUi();
function importAllOrganizationNamesToSheet() {
  var lock = LOCK();
  lock.waitLock(30000); // up to 30s
  try {
    var ss = SS();
    var sheetName = 'Klantnamen ITGlue';
    var sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    // Ensure header row
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, 3).setValues([['Groep', 'Index', 'Naam']]);
    } else {
      // If headers are missing, set them
      var hdrs = sh.getRange(1, 1, 1, 3).getValues()[0];
      if (!hdrs[0] || !hdrs[1] || !hdrs[2]) {
        sh.getRange(1, 1, 1, 3).setValues([['Groep', 'Index', 'Naam']]);
      }
    }

    // Fetch organizations via repo
    var orgs = getOrganizationOverview(); // returns [{id,name}]
    // Sort by name (optional, but nice)
    orgs.sort(function (a, b) { return String(a.name).localeCompare(String(b.name), 'nl'); });

    // Prepare values
    var values = [];
    for (var i = 0; i < orgs.length; i++) {
      values.push(['Klanten', i + 1, orgs[i].name]);
    }

    // Clear old data (rows 2..end, cols A..C)
    var lastRow = sh.getLastRow();
    if (lastRow > 1) {
      sh.getRange(2, 1, lastRow - 1, 3).clearContent();
    }

    if (values.length) {
      sh.getRange(2, 1, values.length, 3).setValues(values);
    }

    ui.alert('Imported ' + values.length + ' organization names from ITGlue into "' + sheetName + '".');
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function compareCustomerNames() {
  var ss = SS();
  var sheetNameKeepass = "Klanten overzicht";
  var sheetNameITGlue = "Klantnamen ITGlue";
  var shk = ss.getSheetByName(sheetNameKeepass);
  var shi = ss.getSheetByName(sheetNameITGlue);

  
}