/**
 * ----------------------------
 * KEEPASS XML IMPORT (→ Google Sheet)
 * ----------------------------
 *
 * Doel:
 * - KeePass XML-bestand (op Google Drive) parsen en relevante entries
 *   naar een Google Sheet schrijven.
 * - Extra helpers om top-level groepen en subgroepen te verkennen.
 *
 * Vereisten:
 * - Toegang tot het XML-bestand op Drive (fileId).
 * - Structuur verwacht door KeePass export:
 *   <KeePassFile><Root> ... <Group><Name>Database</Name> ... </Root></KeePassFile>
 *   Binnen "Database" wordt een groep "Klanten" verwacht met subgroepen per klant.
 *
 * Bladen die gebruikt/aangemaakt worden:
 * - "Klant import"   : resultaat van importKeePassXML()
 * - "Toplevel Groepen": resultaat van listTopGroups()
 * - "Subgroepen"      : resultaat van listSubGroups(index)
 *
 * Let op:
 * - fileId staat hardcoded (pas aan naar jouw bestand).
 * - De XML kan groot zijn: parse en iteratie vinden in memory plaats.
 * - startIndex is 0-based (zoals gevraagd), maxKlanten begrenst het aantal te verwerken klantgroepen.
 */

// ----------------------------
// Hoofdingang: XML inlezen en entries wegschrijven naar "Klant import"
// ----------------------------

// From: KeePassXML.js
function importKeePassXML() {
  var ui = SpreadsheetApp.getUi();

  // Prompt: vanaf welke klant (0-based index)
  var startResp = ui.prompt("XML Import", "Vanaf welke klant (index, begint bij 0)?", ui.ButtonSet.OK_CANCEL);
  if (startResp.getSelectedButton() !== ui.Button.OK) return;
  var startIndex = parseInt(startResp.getResponseText(), 10) || 0;

  // Prompt: maximum aantal klantgroepen om te importeren
  var limitResp = ui.prompt("XML Import", "Hoeveel klanten wil je importeren?", ui.ButtonSet.OK_CANCEL);
  if (limitResp.getSelectedButton() !== ui.Button.OK) return;
  var maxKlanten = parseInt(limitResp.getResponseText(), 10) || 5;

  // ID van het KeePass XML-bestand op Drive (pas aan)
  var fileId = "1jzZ3L8NnecW2cb2IqPyPU6FPvg0KMdXD";

  // Doelblad voorbereiden
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Klant import";
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  sheet.clear();

  // Headers schrijven
  var headers = ["Naam","Username","Wachtwoord","URL","Notities","Path"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // XML-bestand ophalen en als string lezen
  var file = DriveApp.getFileById(fileId);
  var xmlContent = file.getBlob().getDataAsString();
  Logger.log("Bestand opgehaald, lengte: " + xmlContent.length);

  // XML parsen → DOM
  var document = XmlService.parse(xmlContent);
  var root = document.getRootElement(); // <KeePassFile>
  var rootElement = root.getChild("Root");
  if (!rootElement) return; // Geen Root → stop

  // Zoek de "Database" groep direct onder Root
  var databaseGroup = rootElement.getChildren("Group").find(g => g.getChildText("Name") === "Database");
  if (!databaseGroup) return;

  // Zoek de "Klanten" groep binnen "Database"
  var klantenGroup = databaseGroup.getChildren("Group").find(g => g.getChildText("Name") === "Klanten");
  if (!klantenGroup) return;

  // Bepaal het venster van klant-subgroepen dat we willen importeren
  var subGroups = klantenGroup.getChildren("Group").slice(startIndex, startIndex + maxKlanten);

  // Verzamellijst voor rijen die we naar de sheet schrijven
  var rows = [];
  subGroups.forEach(function(klantGroup){
    // Verwerk klantgroep recursief en push rows
    processGroup_Limited(klantGroup, ["Klanten"], rows);
  });

  // Niets gevonden? Stop dan
  if (rows.length === 0) {
    Logger.log("Geen entries gevonden in XML.");
    return;
  }

  // Resultaat in bulk naar sheet
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  Logger.log("Aantal entries geschreven: " + rows.length);
}

// ----------------------------
// Recursieve verwerker: verwerkt subgroepen + entries en bouwt het pad
// ----------------------------

// From: KeePassXML.js
function processGroup_Limited(groupElement, pathArray, rows) {
  // Naam van de huidige groep + breidt pad uit
  var groupName = groupElement.getChildText("Name") || "";
  var currentPath = pathArray.concat([groupName]);

  // 1) Subgroepen eerst recursief verwerken
  var subGroups = groupElement.getChildren("Group");
  subGroups.forEach(function(subG) {
    processGroup_Limited(subG, currentPath, rows);
  });

  // 2) Entries in huidige groep verwerken
  var entries = groupElement.getChildren("Entry");
  entries.forEach(function(entry) {
    // We schrijven exact 6 kolommen: Naam, Username, Wachtwoord, URL, Notities, Path
    var newRow = new Array(6).fill("");

    // Path opmaken als "#Klanten/…/HuidigeGroep[/Subgroep...]"
    newRow[5] = "#" + currentPath.join("/");

    // Key/Value velden van de entry uitlezen
    var strings = entry.getChildren("String");
    strings.forEach(function(s){
      var key = s.getChildText("Key");
      var value = s.getChildText("Value");
      if (key && value) {
        key = key.toLowerCase();
        if (key === "username") newRow[1] = value;
        if (key === "password") newRow[2] = value;
        if (key === "url")      newRow[3] = value;
        if (key === "notes")    newRow[4] = value;
        if (key === "title")    newRow[0] = value;
      }
    });

    // Rij toevoegen aan output
    rows.push(newRow);
  });
}

// ----------------------------
// Hulper om top-level groepen onder "Database" te tonen in een sheet
// ----------------------------

// From: KeePassXML.js
function listTopGroups() {
  var fileId = "1tTTKSodVj-0rUBgdI00vGLmGOngxOa3Y"; // Drive file ID van je KeePass XML

  // Sheet voorbereiden
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Toplevel Groepen";
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  sheet.clear();

  // Headers
  sheet.getRange(1, 1, 1, 2).setValues([["Index", "Toplevel Groep"]]);

  // XML inladen
  var file = DriveApp.getFileById(fileId);
  var xmlContent = file.getBlob().getDataAsString();
  var document = XmlService.parse(xmlContent);
  var root = document.getRootElement();
  var rootElement = root.getChild("Root");
  if (!rootElement) return;

  // Vind "Database"
  var databaseGroup = rootElement.getChildren("Group").find(g => g.getChildText("Name") === "Database");
  if (!databaseGroup) return;

  // Alle top-level groepen onder "Database" verzamelen
  var topGroups = databaseGroup.getChildren("Group");
  var rows = [];
  topGroups.forEach(function(g, idx){
    var name = g.getChildText("Name") || "(geen naam)";
    rows.push([idx, name]);
    Logger.log("Toplevel groep: " + name);
  });

  // Schrijf resultaten
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    Logger.log("Aantal topgroepen: " + rows.length);
  }
}

// ----------------------------
// Hulper om subgroepen van een gekozen top-level groep te tonen
// ----------------------------

// From: KeePassXML.js
function listSubGroups(topGroupIndex) {
  var fileId = "1tTTKSodVj-0rUBgdI00vGLmGOngxOa3Y"; // Drive file ID van je KeePass XML

  // Sheet voorbereiden
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Subgroepen";
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  sheet.clear();

  // Headers
  sheet.getRange(1, 1, 1, 3).setValues([["Index","Naam","Bovenliggende groep"]]);

  // XML inladen
  var file = DriveApp.getFileById(fileId);
  var xmlContent = file.getBlob().getDataAsString();
  var document = XmlService.parse(xmlContent);
  var root = document.getRootElement();
  var rootElement = root.getChild("Root");
  if (!rootElement) return;

  // Vind "Database"
  var databaseGroup = rootElement.getChildren("Group").find(g => g.getChildText("Name") === "Database");
  if (!databaseGroup) return;

  // Controleer geldige index
  var topGroups = databaseGroup.getChildren("Group");
  if (topGroupIndex < 0 || topGroupIndex >= topGroups.length) {
    Logger.log("Ongeldige toplevel index");
    return;
  }

  // Gekozen top-level groep + zijn subgroepen
  var selectedTopGroup = topGroups[topGroupIndex];
  var parentName = selectedTopGroup.getChildText("Name") || "(geen naam)";
  var subGroups = selectedTopGroup.getChildren("Group");

  // Bouw rijen met subgroep-naam + bovenliggende naam
  var rows = [];
  subGroups.forEach(function(g, idx){
    var name = g.getChildText("Name") || "(geen naam)";
    rows.push([idx, name, parentName]);
    Logger.log("Subgroep: " + name + " onder " + parentName);
  });

  // Schrijf resultaten
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    Logger.log("Aantal subgroepen: " + rows.length);
  }
}
