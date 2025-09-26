function myFunction() {
  createPasswordandFolder();
}

function startScript() {
  var ui = SpreadsheetApp.getUi();
  var personResponse = ui.prompt(
    'Wie gebruikt deze optie?',
    '1 voor Corne en 2 voor Kevin',
    ui.ButtonSet.OK_CANCEL
  );
  if (personResponse.getSelectedButton() != ui.Button.OK) {
    ui.alert("Geannuleerd");
    return;
  }
  var personChoice = personResponse.getResponseText().trim();

  if (personChoice !== "1" && personChoice !== "2") {
    ui.alert("Ongeldige invoer, gebruik 1 of 2.");
    return;
  }

  // Stap 2: functie kiezen
  var funcResponse = ui.prompt(
    'Kies een functie',
    'Gebruik 1 voor het formatten van jouw invoer of 2 voor het aanmaken van de wachtwoorden.',
    ui.ButtonSet.OK_CANCEL
  );
  if (funcResponse.getSelectedButton() != ui.Button.OK) {
    ui.alert("Geannuleerd");
    return;
  }
  var funcChoice = funcResponse.getResponseText().trim();

  if (funcChoice !== "1" && funcChoice !== "2") {
    ui.alert("Ongeldige invoer, gebruik 1 of 2.");
    return;
  }

  // Beide keuzes doorgeven aan een centrale functie
  runAction(personChoice, funcChoice);
}

function runAction(person, func) {
  // voorbeeld logica
  if (person === "1") {
    if (func === "1") {
      processPasswords('1');
    } else {
      createPasswordandFolder('1');
    }
  } else if (person === "2") {
    if (func === "1") {
      processPasswords('2');
    } else {
      createPasswordandFolder('2');
    }
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Start')
    .addItem('Start', 'startScript')
    .addItem('Klant import naar input zetten', 'KTIStart')
    .addItem('Test index select', 'testReturn')
    .addToUi();
  ui.createMenu('IT Glue Import Overig')
    .addSeparator()
    .addItem('Import XML', 'importKeePassXML')
    .addToUi();

}