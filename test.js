/**
 * Compare "Klanten overzicht" C (from row 36) against "Klantnamen ITGlue" C (from row 2).
 * Writes ✔ (green) or ✘ (red) into column D on the same rows in "Klanten overzicht".
 * Order does NOT matter.
 */
function markCustomerNameMatches_C36_vs_ITGlueC(opts) {
  opts = opts || {};
  var debug = !!opts.debug;

  // Optional lock (uses your LOCK() if present; otherwise Document lock)
  var lock = (typeof LOCK === 'function') ? LOCK() : LockService.getDocumentLock();
  try { lock.waitLock(30000); } catch(e) {}

  try {
    var ss  = (typeof SS === 'function') ? SS() : SpreadsheetApp.getActiveSpreadsheet();
    var shA = ss.getSheetByName('Klanten overzicht');
    var shB = ss.getSheetByName('Klantnamen ITGlue');
    if (!shA || !shB) throw new Error('Sheets "Klanten overzicht" or "Klantnamen ITGlue" not found.');

    // --- Configuration (as you requested)
    var START_ROW_A = 36; // Klanten overzicht starts at row 36
    var COL_A       = 3;  // C
    var OUT_COL_A   = 4;  // D (output)
    var START_ROW_B = 2;  // Klantnamen ITGlue starts at row 2
    var COL_B       = 3;  // C

    // --- Read reference list (Klantnamen ITGlue!C2:C)
    var lastRowB = shB.getLastRow();
    var namesB = [];
    if (lastRowB >= START_ROW_B) {
      namesB = shB.getRange(START_ROW_B, COL_B, lastRowB - START_ROW_B + 1, 1).getValues().map(function(r){ return r[0]; });
    }

    // Build normalized lookup set
    var lookup = Object.create(null);
    for (var i = 0; i < namesB.length; i++) {
      var key = normCompany(namesB[i]);
      if (key) lookup[key] = true;
    }

    // --- Read working list (Klanten overzicht!C36:C)
    var lastRowA = shA.getLastRow();
    if (lastRowA < START_ROW_A) return; // nothing to do
    var countA = lastRowA - START_ROW_A + 1;
    var namesA = shA.getRange(START_ROW_A, COL_A, countA, 1).getValues().map(function(r){ return r[0]; });

    // --- Compare + mark
    var outMarks  = new Array(namesA.length);
    var outColors = new Array(namesA.length);
    var normDbg   = debug ? new Array(namesA.length) : null;
    var unmatched = debug ? [] : null;

    for (var r = 0; r < namesA.length; r++) {
      var raw = namesA[r];
      if (!raw) {
        outMarks[r]  = [''];
        outColors[r] = [null];
        if (debug) normDbg[r] = [''];
        continue;
      }
      var key = normCompany(raw);
      var found = !!lookup[key];
      outMarks[r]  = [found ? '✔' : '✘'];
      outColors[r] = [found ? '#188038' : '#D93025'];
      if (debug) {
        normDbg[r] = [key];
        if (!found) unmatched.push([String(raw), key]);
      }
    }

    // Write marks to D (same rows)
    var outRng = shA.getRange(START_ROW_A, OUT_COL_A, outMarks.length, 1);
    outRng.setValues(outMarks);
    outRng.setFontColors(outColors);
    outRng.setFontWeights(outMarks.map(function(){ return ['bold']; }));

    // Optional debug output: normalized values + compact report
    if (debug) {
      shA.getRange(START_ROW_A - 1, 5, 1, 1).setValue('Normalized C (from row ' + START_ROW_A + ')');
      shA.getRange(START_ROW_A, 5, normDbg.length, 1).setValues(normDbg);

      var rep = ss.getSheetByName('_Compare Report') || ss.insertSheet('_Compare Report');
      rep.clear();
      rep.getRange(1,1,1,2).setValues([['Unmatched original (Klanten overzicht C from row ' + START_ROW_A + ')','Normalized']]);
      if (unmatched.length) rep.getRange(2,1,unmatched.length,2).setValues(unmatched);

      var sample = Object.keys(lookup).slice(0, 50).map(function(k){ return [k]; });
      rep.getRange(1,4,1,1).setValue('Reference normalized sample (Klantnamen ITGlue C)');
      if (sample.length) rep.getRange(2,4,sample.length,1).setValues(sample);
      rep.autoResizeColumns(1, 5);
    }
  } finally {
    try { lock.releaseLock && lock.releaseLock(); } catch(e) {}
  }
}

/** Strong, case-insensitive normalization for company names. */
function normCompany(v) {
  if (v == null) return '';
  var s = String(v);

  // Normalize whitespace (incl. non-breaking spaces)
  s = s.replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim();

  // Remove diacritics
  if (typeof s.normalize === 'function') {
    s = s.normalize('NFKD').replace(/[\u0300-\u036f]/g, '');
  }

  s = s.toLowerCase().trim();
  s = s.replace(/\s*&\s*/g, ' and ').replace(/\s+/g, ' ');

  // Remove punctuation we don’t care about
  s = s.replace(/[’'".,;:()/\\[\]{}\-–—_]/g, ' ').replace(/\s+/g, ' ').trim();

  // Strip common suffixes at the end (iterate to remove stacked terms)
  var suffixes = [
    'b\\.?v\\.?','n\\.?v\\.?','v\\.?o\\.?f\\.?','c\\.?v\\.?',
    'bvba','gmbh','s\\.?r\\.?l\\.?','ltd','inc\\.?','co\\.?','corp\\.?',
    'holding','group','beheer','invest','investments'
  ];
  var re = new RegExp('\\b(?:' + suffixes.join('|') + ')\\b\\.?$', 'i');
  var prev;
  do { prev = s; s = s.replace(re, '').trim(); } while (s !== prev);

  return s.replace(/\s+/g, ' ').replace(/[.,]+$/g, '').trim();
}

// Convenience wrappers
function markCustomerNameMatches_Run()     { markCustomerNameMatches_C36_vs_ITGlueC({ debug: false }); }
function markCustomerNameMatches_DebugRun(){ markCustomerNameMatches_C36_vs_ITGlueC({ debug: true  }); }



function highlightMatchesInBlad12() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Blad12");

  // Alle data ophalen
  var lastRow = sheet.getLastRow();
  var valuesA = sheet.getRange("A2:A" + lastRow).getValues().flat(); // bron
  var valuesB = sheet.getRange("B2:B" + lastRow).getValues().flat(); // array

  // Reset kleuren en kolom C leegmaken
  sheet.getRange("A2:A" + lastRow).setBackground("white");
  sheet.getRange("C2:C" + lastRow).clearContent();

  // Loop door waarden in kolom A
  valuesA.forEach(function(val, i) {
    if (!val) return;

    // zoek eerste match in kolom B
    var match = valuesB.find(function(cell) {
      return cell && cell.toString().toLowerCase().includes(val.toString().toLowerCase());
    });

    if (match) {
      // Kolom A geel kleuren
      sheet.getRange(i + 2, 1).setBackground("yellow");

      // Matchende waarde in kolom C schrijven
      sheet.getRange(i + 2, 3).setValue(match);
    }
  });
}

