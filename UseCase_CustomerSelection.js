/**
 * ----------------------------
 * CUSTOMER SELECTION & SHEET HELPERS
 * ----------------------------
 *
 * Doel:
 * - Interactieve workflow om een klant (op index) te kiezen, entries uit
 *   "Klant import" te selecteren en over te zetten naar het juiste Input-blad.
 * - Rijen kleuren in "Klanten overzicht" en "Klant import" ter feedback.
 *
 * Verwachte globale helpers/variabelen:
 * - SS():              retornaat het actieve SpreadsheetApp.getActiveSpreadsheet()
 * - LOCK():            retornaat een ScriptLock (LockService) of wrapper met tryLock/waitLock/releaseLock/hasLock
 * - ui:                SpreadsheetApp.getUi() (wordt in KTIStart gezet)
 * - userChoice:        "1" of "2" (wie gebruikt het script)
 *
 * Belangrijke sheets:
 * - "Klanten overzicht" (A=group, B=index, C=name)
 * - "Klant import"      (A..F = entry, user, pass, url, notes, path)
 * - "Input(Corné)" / "Input(Kevin)" voor output
 */

// ----------------------------
// Startstap: vraag gebruiker (Corne/Kevin) en trap workflow af
// ----------------------------

// From: KlantToInput.js
function KTIStart() {
  // UI initialiseren en opslaan in globale 'ui'
  ui = SpreadsheetApp.getUi();

  // Vraag: wie gebruikt het script?
  const res = ui.prompt(
    "Gebruiker",
    "Wie gebruikt het script? 1 = Corne 2 = Kevin",
    ui.ButtonSet.OK_CANCEL
  );

  // Annuleren → stop
  if (res.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Geannuleerd.");
    return;
  }

  // Keuze normaliseren en valideren
  userChoice = (res.getResponseText() || "").trim();
  if (userChoice !== "1" && userChoice !== "2") {
    ui.alert("Ongeldige invoer, gebruik 1 of 2.");
    return;
  }

  // Voer hoofdproces uit. Als dit 'false' retourneert, probeer één keer opnieuw.
  if (KlantToInput() == false) {
    KlantToInput(false);
  }
}

// ----------------------------
// Kies hoe de klant/entries geselecteerd worden
// ----------------------------

// From: KlantToInput.js
function ChooseCustomerIndex() {
  // Prompt: automatische volgende klant of handmatige index
  getEntryChoice = ui.prompt(
    "Hoe wil je de entries krijgen?",
    "1. Krijg entries op basis van eerstvolgende klant in de lijst\n2. Krijg entries op basis van gekozen index.",
    ui.ButtonSet.OK_CANCEL
  );

  if (getEntryChoice.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Geannuleerd.");
    return null;
  }

  // Keuze 1: automatische keuze van eerstvolgende klant
  if (getEntryChoice.getResponseText() === "1") {
    // Probeer direct lock; zo niet, wacht max 30s
    if (!LOCK().tryLock(1)) LOCK().waitLock(30000);

    // selectFirstCustomerEntries() geeft o.a. customerId terug
    if (selectFirstCustomerEntries().customerId) {
      var chosenId = Number(selectFirstCustomerEntries().customerId);
    } else {
      // Foutpad als ID leeg is
      throw new Error(`chosenId was returned empty: "${typeof (chosenId)}": "${chosenId}"`);
    }
    return chosenId;

  } else {
    // Keuze 2: handmatige index
    const res = ui.prompt(
      "Klant index",
      "Wat is de index van de gekozen klant? (zie klantenoverzicht)",
      ui.ButtonSet.OK_CANCEL
    );
    if (res.getSelectedButton() !== ui.Button.OK) {
      ui.alert("Geannuleerd.");
      return null;
    }

    const raw = (res.getResponseText() || "").trim();
    if (raw === "") {
      ui.alert("Geen index ingevoerd.");
      return null;
    }

    // Zorg dat het een integer is
    const ix = Number.parseInt(raw, 10);
    if (Number.isNaN(ix)) {
      ui.alert("Index moet een getal zijn.");
      return null;
    }
    return ix;
  }
}

// ----------------------------
// String normalisatie helpers
// ----------------------------

// From: KlantToInput.js
function normLower(s) {
  // Trim, multiplespaces → 1 spatie, lower-case
  return String(s || "").trim().replace(/\s+/g, " ").toLowerCase();
}

// From: KlantToInput.js
function normalizeName(s) {
  // Trim, lower-case en vervang slashes door spatie (voor vergelijkingen)
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/[\/\\]/g, " "); // replace slashes with space
}

// ----------------------------
// Zoek rijnummer in "Klanten overzicht" op basis van index (alleen group="klanten")
// ----------------------------

// From: KlantToInput.js
function findKlantenRowByIndex(ix) {
  const sh = SS().getSheetByName("Klanten overzicht");
  if (!sh) return null;

  // Data ophalen (header op rij 1)
  const vals = sh.getDataRange().getValues();
  for (let r = 1; r < vals.length; r++) {
    const group = normLower(vals[r][0]);   // Kolom A
    const indexVal = parseInt(vals[r][1], 10); // Kolom B
    if (group === "klanten" && indexVal === ix) {
      // r is 0-based binnen data; +1 voor sheet-rij, +1 voor header → totaal +1
      return r + 1;
    }
  }
  return null;
}

// ----------------------------
// Kleuren van rij in "Klanten overzicht" o.b.v. gekozen index en kleurnaam
// ----------------------------

// From: KlantToInput.js
function ColorRow(inputSheet, rowIndex, chosenColor) {
  if (inputSheet) {
    const targetRow = findKlantenRowByIndex(rowIndex);
    if (targetRow) {
      const lastCol = inputSheet.getLastColumn();
      // Eenvoudige switch op kleurnaam
      switch (chosenColor) {
        case "green":
          inputSheet.getRange(targetRow, 1, 1, lastCol).setBackground("#00ff00");
          break;
        case "red":
          inputSheet.getRange(targetRow, 1, 1, lastCol).setBackground("#ff0000");
          break;
        case "blue":
          inputSheet.getRange(targetRow, 1, 1, lastCol).setBackground("#0000ff");
          break;
        case "orange":
          inputSheet.getRange(targetRow, 1, 1, lastCol).setBackground("#ffa500");
          break;
        default:
          ui.alert("You have chosen a color which was not listed: " + chosenColor);
          break;
      }
    }
  }
}

// ----------------------------
// Kleuren van willekeurige rij in een sheet (hex of naam)
// ----------------------------

// From: KlantToInput.js
function ColorSpecificRow(sheet, rowNumber, hexOrName) {
  if (!sheet || !rowNumber) return;
  const lastCol = sheet.getLastColumn();

  // Map van namen naar hex; anders neem direct gegeven hex of default zachtgeel
  const colors = { green: "#00ff00", red: "#ff0000", blue: "#0000ff", orange: "#ffa500" };
  const color = colors[hexOrName?.toLowerCase?.()] || hexOrName || "#fff9c4";

  sheet.getRange(rowNumber, 1, 1, lastCol).setBackground(color);
}

// ----------------------------
// Hoofdproces: entries ophalen voor gekozen klant en naar Input-blad schrijven
// ----------------------------

// From: KlantToInput.js
function KlantToInput(foundCustomer = true) {
  // Bepaal output-blad op basis van userChoice
  const overzicht = SS().getSheetByName("Klanten overzicht");
  let outputSheet;
  if (userChoice === "1") {
    outputSheet = SS().getSheetByName("Input(Corné)");
  } else if (userChoice === "2") {
    outputSheet = SS().getSheetByName("Input(Kevin)");
  }

  if (!outputSheet) {
    ui.alert("Doelblad niet gevonden.");
    return;
  }

  // Laat gebruiker klant-index kiezen (automatisch of handmatig)
  const chosenIndex = ChooseCustomerIndex();
  if (chosenIndex === null) return;

  // Haal bijpassende entries uit "Klant import"
  json = selectEntriesByCustomerIndex(chosenIndex) || [];
  if (!Array.isArray(json)) {
    ui.alert("Ongeldige data: verwacht een array.");
    return;
  }

  // Headers voor output
  const HEADERS = ['Naam', 'Username', 'Wachtwoord', 'URL', 'Notities', '', 'Path'];

  // Transformeer JSON → rijen
  const rows = json.map(o => ([
    o.entryName ?? '',
    o.username ?? '',
    o.password ?? '',
    o.url ?? '',
    o.notes ?? '',
    '',
    o.path ?? ''
  ]));

  // Output-blad leegmaken en schrijven
  outputSheet.clearContents();
  outputSheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  if (rows.length) {
    // Data wegschrijven
    outputSheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
    // Markeer klant-rij in overzicht als 'gebruikt'
    ColorRow(overzicht, chosenIndex, "green");

  } else if (foundCustomer == false) {
    // Geen data gevonden (tweede poging) → rood markeren en melden
    ColorRow(overzicht, chosenIndex, "red");
    ui.alert("Geen data gevonden voor deze klant. Rij in 'Klanten overzicht' gemarkeerd.");

  } else {
    // Eerste keer geen data → meld en laat KTIStart opnieuw proberen (caller)
    ui.alert("Geen data gevonden voor deze klant. We proberen het nog 1 keer.");
    return false;
  }

  // --- Opschonen van "Klant import": verwijder rijen die we zojuist hebben overgenomen ---
  const shImport = SS().getSheetByName("Klant import");
  const lastRow = shImport.getLastRow();
  if (lastRow > 1) {
    // Lees alle kolommen (A..F) zoals werkelijk aanwezig
    const importLastCol = shImport.getLastColumn();
    const importData = shImport.getRange(2, 1, lastRow - 1, importLastCol).getValues();

    // Bouw keys van net geschreven entries: (entryName + path), allemaal genormaliseerd
    const key = (name, p) => (String(name).trim().toLowerCase() + "||" + String(p).trim().toLowerCase());
    const writtenKeys = new Set(json.map(j => key(j.entryName, j.path)));

    // Zoek rijen in "Klant import" die overeenkomen met geschreven entries
    const rowsToDelete = [];
    for (let r = 0; r < importData.length; r++) {
      const [naam, user, pass, url, notes, path] = importData[r]; // exact 6 kolommen
      if (writtenKeys.has(key(naam, path))) {
        rowsToDelete.push(r + 2); // +2: data start op rij 2
      }
    }

    // Verwijder van onder naar boven om indexverschuiving te voorkomen
    rowsToDelete.reverse().forEach(r => shImport.deleteRow(r));

    // Lock-status tonen en lock vrijgeven (als LOCK() is gebruikt in deze flow)
    ui.alert(LOCK().hasLock());
    LOCK().releaseLock();

    // Terug naar hoofdkeuze-menu
    startScript();
  }
}

// ----------------------------
// Bepaal "eerstvolgende klant": kijk naar eerste Path in "Klant import"
// en verzamel alle entries voor die klant
// ----------------------------

// From: SelectEntriesByCustomer.js
function selectFirstCustomerEntries() {
  const ss = SpreadsheetApp.getActive();
  const SH_IMPORT = 'Klant import';
  const SH_OVERVIEW = 'Klanten overzicht';
  const TARGET_GROUP = 'Klanten';

  // Lees "Klant import"
  const shImport = SS().getSheetByName(SH_IMPORT);
  if (!shImport) throw new Error(`Sheet "${SH_IMPORT}" not found.`);
  const lastRow = shImport.getLastRow();
  const lastCol = shImport.getLastColumn();
  if (lastRow < 2) return { customerName: '', customerId: null, matches: [] };

  // Kolomindexen (1-based voor leesbaarheid; corrigeer met -1 bij data[])
  const COLS = { entry: 1, user: 2, pass: 3, url: 4, notes: 5, path: 6 };
  const data = shImport.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Bepaal klantnaam uit de eerste datarij (Path)
  const firstPath = String((data[0] && data[0][COLS.path - 1]) || '').trim();
  if (!firstPath) throw new Error('First data row has an empty Path in "Klant import".');

  const customerName = extractCustomerFromPath_(firstPath);
  if (!customerName) {
    throw new Error(`Could not extract customer from first Path: "${firstPath}". Expected "#klanten/{customer}[...]".`);
  }

  // Zoek bijbehorend klant-ID (kolom B) in "Klanten overzicht", alleen binnen group "Klanten"
  const customerId = findCustomerIdFromOverview_(ss, SH_OVERVIEW, TARGET_GROUP, customerName);

  // Verzamel alle rijen in "Klant import" die bij deze klant horen
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

// ----------------------------
// Vind klant-ID in "Klanten overzicht" (group=Klanten, name≈customerName)
// ----------------------------

// From: SelectEntriesByCustomer.js
function findCustomerIdFromOverview_(ss, sheetName, targetGroup, customerName) {
  const sh = SS().getSheetByName(sheetName);
  if (!sh) throw new Error(`Sheet "${sheetName}" not found.`);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  // Kolommen: A=group, B=index (ID), C=name
  const rows = sh.getRange(2, 1, lastRow - 1, 3).getValues();

  const matches = [];
  for (const [group, id, name] of rows) {
    // Sla lege rijen over
    if (isBlank_(group) || isBlank_(name)) continue;
    // Alleen group = targetGroup
    if (!equalsLoose_(group, targetGroup)) continue;

    // Naam-vergelijking (genormaliseerd, 'contains' voor tolerant matchen)
    if (normalizeName(name).includes(normalizeName(customerName))) {
      if (id !== null && id !== undefined && String(id).trim() !== '') {
        matches.push(String(id).trim());
      }
    }
  }

  // Geen matches → null
  if (!matches.length) return null;

  // Meerdere matches → neem eerste, log waarschuwing
  if (matches.length > 1) {
    Logger.log(`Warning: multiple IDs found for "${customerName}" in group "${targetGroup}": [${matches.join(', ')}]. Using first.`);
  }
  return matches[0];
}

// ----------------------------
// Kleine helpers (duplicaten van elders voor lokale call sites)
// ----------------------------

// From: SelectEntriesByCustomer.js
function extractCustomerFromPath_(path) {
  // Haal klantnaam uit "#klanten/{naam}/..."
  const m = /^#klanten\/([^\/#]+)(?:\/.*)?$/i.exec(path);
  return m ? m[1].trim() : null;
}

// From: SelectEntriesByCustomer.js
function equalsLoose_(a, b) {
  // Case-insensitive trim-vergelijking
  return String(a).trim().toLowerCase() === String(b).trim().toLowerCase();
}

// From: SelectEntriesByCustomer.js
function isBlank_(v) {
  // Check op null/undefined/lege string na trim
  return v === null || v === undefined || String(v).trim() === '';
}

// ----------------------------
// (Placeholder) testfunctie
// ----------------------------

// From: SelectEntriesByCustomerIndex.js
function testReturn() {
  // Geen implementatie; bedoeld voor debug/uitbreiding
}

// ----------------------------
// Selecteer entries uit "Klant import" op basis van klantindex uit "Klanten overzicht"
// ----------------------------

// From: SelectEntriesByCustomerIndex.js
function selectEntriesByCustomerIndex(chosenIndex) {
  const ss = SpreadsheetApp.getActive();
  const SH_OVERVIEW = 'Klanten overzicht'; // A=group, B=index, C=name
  const SH_IMPORT = 'Klant import';
  const TARGET_GROUP = 'Klanten';

  // Overzichtsheet ophalen
  const shOverview = SS().getSheetByName(SH_OVERVIEW);
  if (!shOverview) throw new Error(`Sheet "${SH_OVERVIEW}" not found.`);

  const lastRowOv = shOverview.getLastRow();
  if (lastRowOv < 2) throw new Error(`No data in "${SH_OVERVIEW}".`);

  // Lees A:C en bouw map index → naam (alleen voor group=Klanten)
  const ovRows = shOverview.getRange(2, 1, lastRowOv - 1, 3).getValues();
  const map = new Map();

  for (const [group, idx, name] of ovRows) {
    if (isBlank_(group) || isBlank_(name)) continue;
    if (String(group).trim().toLowerCase() !== TARGET_GROUP.toLowerCase()) continue;

    if (idx === null || idx === undefined) continue;
    const key = String(idx).trim();
    if (key === '') continue;

    const val = String(name).trim();
    // Als index dubbel voorkomt met andere naam → harde fout
    if (map.has(key) && map.get(key) !== val) {
      throw new Error(`Duplicate index "${key}" with different names in group "${TARGET_GROUP}".`);
    }
    map.set(key, val);
  }

  // Haal klantnaam bij gekozen index
  const customerName = map.get(String(chosenIndex).trim());
  if (!customerName) {
    throw new Error(`No customer found for index ${chosenIndex} in group "${TARGET_GROUP}".`);
  }

  // Lees "Klant import" data
  const shImport = SS().getSheetByName(SH_IMPORT);
  const lastRowIm = shImport.getLastRow();
  const lastColIm = shImport.getLastColumn();
  if (lastRowIm < 2) return [];

  const data = shImport.getRange(2, 1, lastRowIm - 1, lastColIm).getValues();
  const COLS = { entry: 1, user: 2, pass: 3, url: 4, notes: 5, path: 6 };

  const matches = [];
  let pathCustomerForLoose = null;

  // 1) Probeer eerst exact te matchen op klantnaam uit path
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const path = String(row[COLS.path - 1] || '').trim();
    const pathCustomer = extractCustomerFromPath_(path);
    if (!pathCustomer) continue;

    if (equalsLoose_(pathCustomer, customerName)) {
      const sheetRow = r + 2; // data start op rij 2
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
      // Hou een mogelijke deels-match achter de hand
      pathCustomerForLoose = pathCustomer;
    }
  }

  // 2) Geen exacte matches, maar wel mogelijke deelmatch → vraag bevestiging
  if (!matches.length && pathCustomerForLoose &&
    normalizeName(customerName).includes(normalizeName(pathCustomerForLoose))) {

    const partMatch = ui.prompt(
      "Deelse match",
      `Klantnaam uit path: "${pathCustomerForLoose}"\nKlantnaam uit overzicht: "${customerName}"\n\nAccepteren?`,
      ui.ButtonSet.OK_CANCEL
    );

    if (partMatch.getSelectedButton() === ui.Button.OK) {
      // Verzamel alle deelmatches
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

  // 3) Kleur alle te gebruiken rijen in "Klant import" (visuele feedback)
  if (matches.length) {
    matches.forEach(m => ColorSpecificRow(shImport, m.sheetRow, "green"));
  }

  Logger.log(`Group "${TARGET_GROUP}", index ${chosenIndex} → "${customerName}", matches: ${matches.length}`);
  return matches;

  // Lokale helpers (shadowen globale versies indien nodig)
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

// ----------------------------
// Duplicaat helpers (beschikbaar voor andere modules)
// ----------------------------

// From: SelectEntriesByCustomerIndex.js
function isBlank_(v) {
  return v === null || v === undefined || String(v).trim() === '';
}

// From: SelectEntriesByCustomerIndex.js
function extractCustomerFromPath_(path) {
  const m = /^#klanten\/([^\/#]+)(?:\/.*)?$/i.exec(path);
  return m ? m[1].trim() : null;
}

// From: SelectEntriesByCustomerIndex.js
function equalsLoose_(a, b) {
  return String(a).trim().toLowerCase() === String(b).trim().toLowerCase();
}

// From: SelectEntriesByCustomerIndex.js
function normalizeName(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/[\/\\]/g, " ")
    .replace(/\s+/g, " ");
}
