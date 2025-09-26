function importKeePassXML() {
  
  var ui = SpreadsheetApp.getUi();

  // Vraag vanaf welke klant (index)
  var startResp = ui.prompt("XML Import", "Vanaf welke klant (index, begint bij 0)?", ui.ButtonSet.OK_CANCEL);
  if (startResp.getSelectedButton() !== ui.Button.OK) return;
  var startIndex = parseInt(startResp.getResponseText(), 10) || 0;

  // Vraag aantal klanten
  var limitResp = ui.prompt("XML Import", "Hoeveel klanten wil je importeren?", ui.ButtonSet.OK_CANCEL);
  if (limitResp.getSelectedButton() !== ui.Button.OK) return;
  var maxKlanten = parseInt(limitResp.getResponseText(), 10) || 5;

  var fileId = "1tTTKSodVj-0rUBgdI00vGLmGOngxOa3Y"; // Zet hier je Drive file ID
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Klant import";
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  sheet.clear();

  // Headers
  var headers = ["Naam","Username","Wachtwoord","URL","Notities","Path"];
  sheet.getRange(1,1,1,headers.length).setValues([headers]);

  // Haal XML op
  var file = DriveApp.getFileById(fileId);
  var xmlContent = file.getBlob().getDataAsString();
  Logger.log("Bestand opgehaald, lengte: " + xmlContent.length);

  // Parse XML
  var document = XmlService.parse(xmlContent);
  var root = document.getRootElement(); // <KeePassFile>
  var rootElement = root.getChild("Root");
  if (!rootElement) return;

  // Zoek Database-groep
  var databaseGroup = rootElement.getChildren("Group").find(g => g.getChildText("Name") === "Database");
  if (!databaseGroup) return;

  // Zoek Klanten-groep
  var klantenGroup = databaseGroup.getChildren("Group").find(g => g.getChildText("Name") === "Klanten");
  if (!klantenGroup) return;

   // Beperk tot N klanten (met offset)
  var subGroups = klantenGroup.getChildren("Group").slice(startIndex, startIndex + maxKlanten);

  var rows = [];
  subGroups.forEach(function(klantGroup){
    processGroup_Limited(klantGroup, ["Klanten"], rows);
  });

  if (rows.length === 0) {
    Logger.log("Geen entries gevonden in XML.");
    return;
  }

  // Schrijf naar sheet
  sheet.getRange(2,1,rows.length,headers.length).setValues(rows);
  Logger.log("Aantal entries geschreven: " + rows.length);
}


// Verwerkt groepen, beperkt tot geselecteerde kolommen
function processGroup_Limited(groupElement, pathArray, rows) {
  var groupName = groupElement.getChildText("Name") || "";
  var currentPath = pathArray.concat([groupName]);

  // Verwerk subgroepen
  var subGroups = groupElement.getChildren("Group");
  subGroups.forEach(function(subG) {
    processGroup_Limited(subG, currentPath, rows);
  });

  // Verwerk entries
  var entries = groupElement.getChildren("Entry");
  entries.forEach(function(entry) {
    var newRow = new Array(6).fill(""); // Alleen Naam, Username, Wachtwoord, URL, Notities, Path

   
     

    // Path
    newRow[5] = "#" + currentPath.join("/");

    // Key/Value velden
    var strings = entry.getChildren("String");
    strings.forEach(function(s){
      var key = s.getChildText("Key");
      var value = s.getChildText("Value");
      if (key && value) {
        key = key.toLowerCase();
        if (key === "username") newRow[1] = value;
        if (key === "password") newRow[2] = value;
        if (key === "url") newRow[3] = value;
        if (key === "notes") newRow[4] = value;
        if (key === "title") newRow[0] = value;
      }
    });

    rows.push(newRow);
  });
}



function listTopGroups() {
  var fileId = "1tTTKSodVj-0rUBgdI00vGLmGOngxOa3Y"; // Drive file ID van je KeePass XML
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Toplevel Groepen";
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  sheet.clear();

  sheet.getRange(1,1,1,2).setValues([["Index", "Toplevel Groep"]]);

  var file = DriveApp.getFileById(fileId);
  var xmlContent = file.getBlob().getDataAsString();
  var document = XmlService.parse(xmlContent);
  var root = document.getRootElement();
  var rootElement = root.getChild("Root");
  if (!rootElement) return;

  var databaseGroup = rootElement.getChildren("Group").find(g => g.getChildText("Name") === "Database");
  if (!databaseGroup) return;

  var topGroups = databaseGroup.getChildren("Group");
  var rows = [];
  topGroups.forEach(function(g, idx){
    var name = g.getChildText("Name") || "(geen naam)";
    rows.push([idx, name]);
    Logger.log("Toplevel groep: " + name);
  });

  if (rows.length > 0) {
    sheet.getRange(2,1,rows.length,2).setValues(rows);
    Logger.log("Aantal topgroepen: " + rows.length);
  }
}

// Toont subgroepen voor een geselecteerde toplevel-groep
function listSubGroups(topGroupIndex) {
  var fileId = "1tTTKSodVj-0rUBgdI00vGLmGOngxOa3Y"; // Drive file ID van je KeePass XML
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Subgroepen";
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  sheet.clear();

  sheet.getRange(1,1,1,3).setValues([["Index","Naam","Bovenliggende groep"]]);

  var file = DriveApp.getFileById(fileId);
  var xmlContent = file.getBlob().getDataAsString();
  var document = XmlService.parse(xmlContent);
  var root = document.getRootElement();
  var rootElement = root.getChild("Root");
  if (!rootElement) return;

  var databaseGroup = rootElement.getChildren("Group").find(g => g.getChildText("Name") === "Database");
  if (!databaseGroup) return;

  var topGroups = databaseGroup.getChildren("Group");
  if (topGroupIndex < 0 || topGroupIndex >= topGroups.length) {
    Logger.log("Ongeldige toplevel index");
    return;
  }

  var selectedTopGroup = topGroups[topGroupIndex];
  var parentName = selectedTopGroup.getChildText("Name") || "(geen naam)";
  var subGroups = selectedTopGroup.getChildren("Group");
  var rows = [];
  subGroups.forEach(function(g, idx){
    var name = g.getChildText("Name") || "(geen naam)";
    rows.push([idx, name, parentName]);
    Logger.log("Subgroep: " + name + " onder " + parentName);
  });

  if (rows.length > 0) {
    sheet.getRange(2,1,rows.length,3).setValues(rows);
    Logger.log("Aantal subgroepen: " + rows.length);
  }
}