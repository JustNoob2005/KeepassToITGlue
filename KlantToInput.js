const ss = SpreadsheetApp.getActiveSpreadsheet();
let ui;
let userChoice;
let getEntryChoice;
let json;
var lock = LockService.getDocumentLock();

function KTIStart() {
  ui = SpreadsheetApp.getUi();

  const res = ui.prompt(
    "Gebruiker",
    "Wie gebruikt het script? 1 = Corne 2 = Kevin",
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Geannuleerd.");
    return;
  }

  userChoice = (res.getResponseText() || "").trim();
  if (userChoice !== "1" && userChoice !== "2") {
    ui.alert("Ongeldige invoer, gebruik 1 of 2.");
    return;
  }

  if (KlantToInput() == false) {
    KlantToInput(false);
  };

}

function ChooseCustomerIndex() {
  getEntryChoice = ui.prompt(
    "Hoe wil je de entries krijgen?",
    "1. Krijg entries op basis van eerstvolgende klant in de lijst\n2. Krijg entries op basis van gekozen index.",
    ui.ButtonSet.OK_CANCEL
  )

  if (getEntryChoice.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Geannuleerd.");
    return null;
  }

  if (getEntryChoice.getResponseText() === "1") {
    if (!lock.tryLock(1)) lock.waitLock(30000);
    if (selectFirstCustomerEntries().customerId) {

      var chosenId = Number(selectFirstCustomerEntries().customerId);
    }
    else {
      throw new Error(`chosenId was returned empty: "${typeof (chosenId)}": "${chosenId}"`)
    }
    return chosenId;
  } else {
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

    const ix = Number.parseInt(raw, 10);
    if (Number.isNaN(ix)) {
      ui.alert("Index moet een getal zijn.");
      return null;
    }
    return ix;
  }
}

function normLower(s) {
  return String(s || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function normalizeName(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/[\/\\]/g, " "); // replace slashes with space
}

// Find the row index (1-based) in "Klanten overzicht" for a given Klanten-group index
function findKlantenRowByIndex(ix) {
  const sh = ss.getSheetByName("Klanten overzicht");
  if (!sh) return null;

  const vals = sh.getDataRange().getValues(); // assumes headers on row 1
  for (let r = 1; r < vals.length; r++) {
    const group = normLower(vals[r][0]); // Col A
    const indexVal = parseInt(vals[r][1], 10); // Col B
    if (group === "klanten" && indexVal === ix) {
      return r + 1; // convert 0-based data row -> 1-based sheet row
    }
  }
  return null;
}

function ColorRow(inputSheet, rowIndex, chosenColor) {
  if (inputSheet) {
    const targetRow = findKlantenRowByIndex(rowIndex);
    if (targetRow) {
      const lastCol = inputSheet.getLastColumn();
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

function ColorSpecificRow(sheet, rowNumber, hexOrName) {
  if (!sheet || !rowNumber) return;
  const lastCol = sheet.getLastColumn();
  // optional: map names to hex if you want to reuse "red/green/blue/orange"
  const colors = { green: "#00ff00", red: "#ff0000", blue: "#0000ff", orange: "#ffa500" };
  const color = colors[hexOrName?.toLowerCase?.()] || hexOrName || "#fff9c4";
  sheet.getRange(rowNumber, 1, 1, lastCol).setBackground(color);
}


// transfer the data to the input sheet for the selected user.
function KlantToInput(foundCustomer = true) {
  const overzicht = ss.getSheetByName("Klanten overzicht");
  let outputSheet;
  if (userChoice === "1") {
    outputSheet = ss.getSheetByName("Input(Corné)");
  } else if (userChoice === "2") {
    outputSheet = ss.getSheetByName("Input(Kevin)");
  }

  if (!outputSheet) {
    ui.alert("Doelblad niet gevonden.");
    return;
  }

  const chosenIndex = ChooseCustomerIndex();
  if (chosenIndex === null) return;
  json = selectEntriesByCustomerIndex(chosenIndex) || [];
  if (!Array.isArray(json)) {
    ui.alert("Ongeldige data: verwacht een array.");
    return;
  }

  const HEADERS = ['Naam', 'Username', 'Wachtwoord', 'URL', 'Notities', '', 'Path'];

  const rows = json.map(o => ([
    o.entryName ?? '',
    o.username ?? '',
    o.password ?? '',
    o.url ?? '',
    o.notes ?? '',
    '',
    o.path ?? ''
  ]));

  // Write
  outputSheet.clearContents();
  outputSheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  if (rows.length) {
    outputSheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
    ColorRow(overzicht, chosenIndex, "green");
  }
  else if (foundCustomer == false) {
    ColorRow(overzicht, chosenIndex, "red")
    ui.alert("Geen data gevonden voor deze klant. Rij in 'Klanten overzicht' gemarkeerd.");
  }
  else {
    ui.alert("Geen data gevonden voor deze klant. We proberen het nog 1 keer.");
    return false;
  }

  const shImport = ss.getSheetByName("Klant import");
  const lastRow = shImport.getLastRow();
  if (lastRow > 1) {
    // Read the REAL columns of Klant import (6 cols, no spacer)
    const importLastCol = shImport.getLastColumn(); // should be 6
    const importData = shImport.getRange(2, 1, lastRow - 1, importLastCol).getValues();

    // Build a fast lookup of what we just wrote (entryName + path), normalized
    const key = (name, p) => (String(name).trim().toLowerCase() + "||" + String(p).trim().toLowerCase());
    const writtenKeys = new Set(json.map(j => key(j.entryName, j.path)));

    const rowsToDelete = [];
    for (let r = 0; r < importData.length; r++) {
      // Destructure EXACTLY the 6 import columns (no spacer)
      const [naam, user, pass, url, notes, path] = importData[r];
      if (writtenKeys.has(key(naam, path))) {
        rowsToDelete.push(r + 2); // +2 because data starts on row 2
      }
    }

    // Debug if needed
    // Logger.log("rowsToDelete: %s", JSON.stringify(rowsToDelete));

    // Delete bottom-up
    rowsToDelete.reverse().forEach(r => shImport.deleteRow(r));
    ui.alert(lock.hasLock());
    lock.releaseLock();
    startScript();
  }
}