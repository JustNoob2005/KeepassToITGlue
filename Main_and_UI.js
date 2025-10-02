// ----------------------------
// Entry points en menu/UI wiring
// ----------------------------

/** Auto-generated restructure. */
// Functie waarmee de gebruiker eerst kiest wie het script gebruikt
// (Corne of Kevin) en daarna welke actie uitgevoerd moet worden
function startScript() {
  // Stap 1: Vraag via een prompt wie de gebruiker is
  var ui = SpreadsheetApp.getUi();
  var personResponse = ui.prompt(
    'Wie gebruikt deze optie?',
    '1 voor Corne en 2 voor Kevin',
    ui.ButtonSet.OK_CANCEL
  );

  // Als de gebruiker op Annuleren klikt, stop de functie
  if (personResponse.getSelectedButton() != ui.Button.OK) {
    ui.alert("Geannuleerd");
    return;
  }

  // Lees keuze van de gebruiker uit (1 of 2)
  var personChoice = personResponse.getResponseText().trim();

  // Controleer of de keuze geldig is (alleen 1 of 2 toegestaan)
  if (personChoice !== "1" && personChoice !== "2") {
    ui.alert("Ongeldige invoer, gebruik 1 of 2.");
    return;
  }

  // Stap 2: Vraag welke functie uitgevoerd moet worden
  var funcResponse = ui.prompt(
    'Kies een functie',
    'Gebruik 1 voor het formatten van jouw invoer of 2 voor het aanmaken van de wachtwoorden.',
    ui.ButtonSet.OK_CANCEL
  );

  // Als de gebruiker op Annuleren klikt, stop de functie
  if (funcResponse.getSelectedButton() != ui.Button.OK) {
    ui.alert("Geannuleerd");
    return;
  }

  // Lees functie-keuze uit (1 of 2)
  var funcChoice = funcResponse.getResponseText().trim();

  // Controleer of de keuze geldig is (alleen 1 of 2 toegestaan)
  if (funcChoice !== "1" && funcChoice !== "2") {
    ui.alert("Ongeldige invoer, gebruik 1 of 2.");
    return;
  }

  // Beide keuzes (persoon en functie) doorgeven aan centrale afhandel-functie
  runAction(personChoice, funcChoice);
}

// ----------------------------
// Run actie voor gekozen gebruiker
// ----------------------------

// Voert de juiste functie uit op basis van wie de gebruiker is en welke functie is gekozen
function runAction(person, func) {
  if (person === "1") { // Als Corne gekozen is
    if (func === "1") {
      processPasswords('1'); // Voer "processPasswords" uit voor Corne
    } else {
      createPasswordandFolder('1'); // Voer "createPasswordandFolder" uit voor Corne
    }
  } else if (person === "2") { // Als Kevin gekozen is
    if (func === "1") {
      processPasswords('2'); // Voer "processPasswords" uit voor Kevin
    } else {
      createPasswordandFolder('2'); // Voer "createPasswordandFolder" uit voor Kevin
    }
  }
}

// ----------------------------
// UI-menu toevoegen bij openen van spreadsheet
// ----------------------------

// Wordt automatisch uitgevoerd bij openen van de spreadsheet
// Maakt menu’s aan in Google Sheets bovenin de menubalk
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // Menu "Start" met 3 opties
  ui.createMenu('Start')
    .addItem('Start', 'startScript') // Roept startScript() aan
    .addItem('Klant import naar input zetten', 'KTIStart') // Roept KTIStart() aan
    .addItem('Test index select', 'testReturn') // Roept testReturn() aan
    .addToUi();

  // Menu "IT Glue Import Overig" met scheiding en 1 optie
  ui.createMenu('IT Glue Import Overig')
    .addSeparator()
    .addItem('Import XML', 'importKeePassXML') // Roept importKeePassXML() aan
    .addItem('Import organization names', 'importAllOrganizationNamesToSheet')
    .addToUi();
}

// ----------------------------
// Hulpfuncties
// ----------------------------

// Controleer of een waarde leeg is (null, undefined of lege string)
function isBlank_(v) {
  return v === null || v === undefined || String(v).trim() === '';
}

// Haalt klantnaam uit een pad zoals "#klanten/Bedrijfsnaam/..."
// Regex: neem alles tussen "#klanten/" en de eerste volgende "/" (indien aanwezig)
// Voorbeeld: "#klanten/Aerson B.V./..." → "Aerson B.V."
function extractCustomerFromPath_(path) {
  const m = /^#klanten\/([^\/#]+)(?:\/.*)?$/i.exec(path);
  return m ? m[1].trim() : null;
}

// V
