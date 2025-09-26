function createPasswordandFolder(person) {

var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
var range = sheet.getDataRange(); 
var values = range.getValues();

  //rows ophalen
  var headers = values[0];
  var organizationNameIndex = headers.indexOf("organization-name");
  var organizationIdIndex = headers.indexOf("organization-id");
  var nameIndex = headers.indexOf("name");
  var usernameIndex = headers.indexOf("username");
  var passwordCategoryIdIndex = headers.indexOf("password-category-id");
  var passwordIndex = headers.indexOf("password");
  var folderLevel1Index = headers.indexOf("MapNiveau1");
  var folderLevel2Index = headers.indexOf("MapNiveau2");
  var folderLevel3Index = headers.indexOf("MapNiveau3");
  var folderLevel4Index = headers.indexOf("MapNiveau4");
  var urlIndex = headers.indexOf("url");
  var notesIndex = headers.indexOf("notes");

  // voor elke row met values uitvoeren
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var organizationName = row[organizationNameIndex];
    var retrievedOrganizationId = getOrganizationID(organizationName);
    var organizationId = retrievedOrganizationId || row[organizationIdIndex];

    //Het organisatie ID ophalen voor deze klant
    if (retrievedOrganizationId) {
      sheet.getRange(i + 1, organizationIdIndex + 1).setValue(retrievedOrganizationId);
    }

    // cellen ophalen
    var name = row[nameIndex];
    var username = row[usernameIndex];
    var passwordCategoryId = row[passwordCategoryIdIndex];
    var password = row[passwordIndex];
    var folderNames= []
    var parentId = null;
    var folderLevel1Name = row[folderLevel1Index] ? row[folderLevel1Index].trim() : "";
    var folderLevel2Name = row[folderLevel2Index] ? row[folderLevel2Index].trim() : "";
    var folderLevel3Name = row[folderLevel3Index] ? row[folderLevel3Index].trim() : "";
    var folderLevel4Name = row[folderLevel4Index] ? row[folderLevel4Index].trim() : "";
  

   
    if(folderLevel1Name){
    folderNames.push(folderLevel1Name)}
     if(folderLevel2Name){
    folderNames.push(folderLevel2Name)}
        if(folderLevel3Name){
    folderNames.push(folderLevel3Name)}
      if(folderLevel4Name){
    folderNames.push(folderLevel4Name)}

    //Logger.log(folderNames);

    var passurl = row[urlIndex];
    var notes = row[notesIndex];

  folderNames.forEach((folder, index) => {
    if (!folder) return;

       var folderId = findPasswordFolder(organizationId,folder,parentId); //ophalen van het mogelijke folder ID voor de organisatie met dit ID, de betreffende folder en met het parentId.
       

    //Logger.log(`Het gevonden Folder id is: ${folderId}`);

    Logger.log("Parent ID:"+parentId)
    Logger.log("Folder ID:"+folderId)

    if (!folderId) {
    //  Logger.log(`Map ${folder} bestaat niet. Aanmaken...`);
      folderId = createPasswordFolder(organizationId, folder,parentId);
     // Logger.log(folderId);
    }

     parentId = folderId;

  });
    var folderId = parentId || null

    if (checkIfPasswordExists(name, organizationId, folderId)) {
      // Kleur de rij lichtblauw
      Logger.log("Het wachtwoord bestaat al!")
      sheet.getRange(i + 1, 1, 1, sheet.getLastColumn())
        .setBackground("#ADD8E6"); // Lichtblauwe kleur
        sheet.getRange(i + 1, 17).setValue(true);
    } else {
      // Maak het wachtwoord aan
      var result = createPassword(organizationId, name, username, passwordCategoryId, password, folderId, notes, passurl);
      if (result) {
        Logger.log(`Wachtwoord succesvol aangemaakt in map "${folderId || 'Hoofdmap'}".`);
        sheet.getRange(i + 1, 17).setValue(true);
      } else {
        Logger.log('Fout bij het aanmaken van het wachtwoord.');
      } 
    }
  }
}

function checkIfPasswordExists(name, organizationId, passwordfolderId) {

  // Haal de wachtwoorden van de organisatie op
  var passwords = getPasswordsFromOrganization(organizationId);

  // Filter de wachtwoorden die aan de voorwaarden voldoen
  var matchingPasswords = passwords.filter(function(password) {

    var nameMatches = String(password.attributes.name) === String(name);
    var folderMatches = !passwordfolderId || String(password.attributes['password-folder-id']) === String(passwordfolderId);
    Logger.log(folderMatches)
    return nameMatches && folderMatches;
  });

  if (matchingPasswords.length > 0) {
    Logger.log('Er zijn ' + matchingPasswords.length + ' wachtwoorden gevonden met de gegeven naam en folder-id.');
    matchingPasswords.forEach(function(password) {
      Logger.log('Wachtwoord ID: ' + password.id);
      Logger.log('Folder ID: ' + password.attributes['password-folder-id']);
    });
    return matchingPasswords;
  } else {
    var folderMessage = passwordfolderId ? ` en folder-id "${passwordfolderId}"` : '';
    Logger.log('Geen wachtwoorden gevonden met de naam "' + name + '"' + folderMessage + ' voor organisatie ' + organizationId);
    return;
  }
}


function getSheetByPerson(person) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName;

  switch (person) {
    case "1":
      sheetName = "Import(Corné)";
      break;
    case "2":
      sheetName = "Import(Kevin)";
      break;
    default:
      throw new Error("Onbekende persoon: " + person);
  }

  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error("Tabblad niet gevonden: " + sheetName);
  }

  return sheet;
}

/**
 * Haalt alle data vanaf A2 tot laatste rij/kolom op
 */
function getDataFromPersonSheet(person) {
  var sheet = getSheetByPerson(person);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2) return []; // Geen data onder header

  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}