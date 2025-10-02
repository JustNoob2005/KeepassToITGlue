/**
 * ----------------------------
 * PASSWORD REPOSITORY & IMPORT HELPERS (Sheet → ITGlue)
 * ----------------------------
 *
 * Doel:
 * - Helpers om wachtwoorden en mappen te lezen/schrijven tussen Google Sheets en ITGlue.
 * - Omvat: aanmaken van wachtwoorden, mappen aanmaken/opzoeken, exporteren naar sheet,
 *   normaliseren/valideren van gegevens, en API-calls naar ITGlue.
 *
 * Vereisten:
 * - Script Properties:
 *   - IT_GLUE_API_KEY (voor standaard API-calls met x-api-key header)
 *   - IT_GLUE_SESSION_TOKEN (voor endpoints die Authorization header vereisen)
 *
 * Belangrijkste functies:
 * - createPasswordandFolder(person): leest rijen uit de actieve sheet, zorgt voor mapstructuur,
 *   checkt of wachtwoord al bestaat, en maakt zo nodig een nieuw wachtwoord in ITGlue.
 * - checkIfPasswordExists(...): controleert of een wachtwoord al bestaat in (sub)map.
 * - getPasswordsFromOrganization(organizationId): haalt alle wachtwoorden van een organisatie (met paginatie).
 * - createPassword(...): maakt een wachtwoord aan in ITGlue.
 * - createPasswordFolder(...), findPasswordFolder(...): map aanmaken/zoeken (met parent).
 * - processPasswords(person): transformeert inputblad → importblad + map-velden, voegt https toe, valideert/kleurt.
 * - writePasswordsToSheet(...): schrijft opgehaalde wachtwoorden naar een sheet.
 * - exportPasswordsForOrganization(...): haalt wachtwoorden op voor een organisatie en exporteert naar sheet.
 * - Hulpfuncties voor per-persoon bladen, caching, URL-reparatie, validatie/kleuren, en sheets legen.
 */

// ----------------------------
// Import: wachtwoorden aanmaken + mapstructuur opbouwen
// ----------------------------

// From: Import passwords.js
function createPasswordandFolder(person) {
  // Actieve sheet + alle waarden ophalen
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();

  // Header-rij ophalen en kolomindexen bepalen (veilig t.o.v. kolomvolgorde)
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

  // Loop over alle data-rijen (vanaf i=1; i=0 is header)
  for (var i = 1; i < values.length; i++) {
    var row = values[i];

    // Klantnaam → organisatie-ID opzoeken, of bestaande ID uit sheet gebruiken
    var organizationName = row[organizationNameIndex];
    var retrievedOrganizationId = getOrganizationID(organizationName); // API lookup
    var organizationId = retrievedOrganizationId || row[organizationIdIndex];

    // Als we via API een ID vonden, schrijf die terug in de sheet (kolom "organization-id")
    if (retrievedOrganizationId) {
      sheet.getRange(i + 1, organizationIdIndex + 1).setValue(retrievedOrganizationId);
    }

    // Basisvelden uit de rij
    var name = row[nameIndex];
    var username = row[usernameIndex];
    var passwordCategoryId = row[passwordCategoryIdIndex];
    var password = row[passwordIndex];

    // Mapniveaus (lege strings als niet gevuld) en lijst voor hiërarchie
    var folderNames = [];
    var parentId = null; // parent-id voor geneste mappen
    var folderLevel1Name = row[folderLevel1Index] ? row[folderLevel1Index].trim() : "";
    var folderLevel2Name = row[folderLevel2Index] ? row[folderLevel2Index].trim() : "";
    var folderLevel3Name = row[folderLevel3Index] ? row[folderLevel3Index].trim() : "";
    var folderLevel4Name = row[folderLevel4Index] ? row[folderLevel4Index].trim() : "";

    // Voeg enkel ingevulde niveaus toe in volgorde
    if (folderLevel1Name) folderNames.push(folderLevel1Name);
    if (folderLevel2Name) folderNames.push(folderLevel2Name);
    if (folderLevel3Name) folderNames.push(folderLevel3Name);
    if (folderLevel4Name) folderNames.push(folderLevel4Name);

    var passurl = row[urlIndex];
    var notes = row[notesIndex];

    // Voor elke mapnaam in de hiërarchie: bestaand map-ID zoeken, zo niet → aanmaken
    folderNames.forEach((folder, index) => {
      if (!folder) return;

      // Probeer bestaande map te vinden op dit niveau (met parentId)
      var folderId = findPasswordFolder(organizationId, folder, parentId);

      // Indien niet gevonden → map aanmaken onder current parentId
      if (!folderId) {
        folderId = createPasswordFolder(organizationId, folder, parentId);
      }

      // Volgende niveau krijgt deze map als parent
      parentId = folderId;
    });

    // Eindmap = diepste gevonden/nieuw aangemaakte map
    var folderId = parentId || null;

    // Bestaat wachtwoord met deze naam + (optioneel) map al?
    if (checkIfPasswordExists(name, organizationId, folderId)) {
      // Bestaat al → kleur hele rij lichtblauw + markeer in kolom Q (17)
      Logger.log("Het wachtwoord bestaat al!");
      sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()).setBackground("#ADD8E6");
      sheet.getRange(i + 1, 17).setValue(true);
    } else {
      // Nog niet bestaand → aanmaken via API
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

// ----------------------------
// Check of wachtwoord al bestaat (op naam + optioneel map)
// ----------------------------

// From: Import passwords.js
function checkIfPasswordExists(name, organizationId, passwordfolderId) {
  // Haal alle wachtwoorden voor de organisatie op (kan groot zijn; cache waar mogelijk)
  var passwords = getPasswordsFromOrganization(organizationId);

  // Filter op exacte naam + (optioneel) zelfde map
  var matchingPasswords = passwords.filter(function (password) {
    var nameMatches = String(password.attributes.name) === String(name);
    var folderMatches = !passwordfolderId || String(password.attributes['password-folder-id']) === String(passwordfolderId);
    Logger.log(folderMatches);
    return nameMatches && folderMatches;
  });

  // Matches → log + return array; anders null/undefined
  if (matchingPasswords.length > 0) {
    Logger.log('Er zijn ' + matchingPasswords.length + ' wachtwoorden gevonden met de gegeven naam en folder-id.');
    matchingPasswords.forEach(function (password) {
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

// ----------------------------
// Hulpfuncties voor per-persoon sheets
// ----------------------------

// From: Import passwords.js
function getSheetByPerson(person) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName;

  // Kies tabblad op basis van persoon (1 = Corné, 2 = Kevin)
  switch (person) {
    case "1":
      sheetName = "Import(Corné)";
      break;
    case "2":
      sheetName = "Import(Kevin)";
      break;
    default:
      // Onbekende code → gooi duidelijke fout
      throw new Error("Onbekende persoon: " + person);
  }

  // Haal sheet op; als ontbreekt → fout
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error("Tabblad niet gevonden: " + sheetName);
  }

  return sheet;
}

// From: Import passwords.js
function getDataFromPersonSheet(person) {
  // Haal de juiste sheet op
  var sheet = getSheetByPerson(person);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  // Geen data (alleen header) → lege array
  if (lastRow < 2) return [];

  // Retourneer alle data onder de header
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

// ----------------------------
// API-calls: Passwords ophalen voor organisatie (met paginatie)
// ----------------------------

// From: PasswordAPICalls.js
function getPasswordsFromOrganization(organizationId) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var apiKey = scriptProperties.getProperty('IT_GLUE_API_KEY');

  // Zonder API key stoppen we (lege lijst)
  if (!apiKey) {
    Logger.log('API Key is missing');
    return [];
  }

  // Eerste pagina = 1 (ITGlue pagineert)
  var allPasswords = [];

  try {
    var hasMorePages = true;
    var currentPage = 1;

    // Zolang 'links.next' aanwezig is → volgende pagina ophalen
    while (hasMorePages) {
      var options = {
        'method': 'GET',
        'headers': {
          'Content-Type': 'application/vnd.api+json',
          'x-api-key': apiKey
        },
        'muteHttpExceptions': true
      };

      var pageUrl = `https://api.eu.itglue.com/passwords?filter[organization_id]=${organizationId}&page[number]=${currentPage}`;

      var response = UrlFetchApp.fetch(pageUrl, options);
      var responseCode = response.getResponseCode();

      if (responseCode !== 200) {
        // Fout → breek af met lege lijst
        return [];
      }

      var responseText = response.getContentText();
      if (!responseText || responseText.trim() === '') {
        return [];
      }

      var data = JSON.parse(responseText);
      var passwords = data.data || [];

      // Voeg toe aan totaal
      allPasswords = allPasswords.concat(passwords);

      // Volgende pagina?
      hasMorePages = data.links && data.links.next;
      currentPage++;
    }

    return allPasswords;

  } catch (error) {
    Logger.log('Error fetching passwords from IT Glue API: ' + error);
    return [];
  }
}

// ----------------------------
// API-call: Specifiek wachtwoord ophalen op ID (relatie-endpoint)
// ----------------------------

// From: PasswordAPICalls.js
function getPasswordById(organizationId, passwordId) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var apiKey = scriptProperties.getProperty('IT_GLUE_API_KEY');

  if (!apiKey) {
    Logger.log('API Key is missing');
    return [];
  }

  // Endpoint voor relatie: organization → passwords → {id}
  var url = `https://api.eu.itglue.com/organizations/${organizationId}/relationships/passwords/${passwordId}`;

  var options = {
    'method': 'GET',
    'headers': {
      'Content-Type': 'application/vnd.api+json',
      'x-api-key': apiKey
    },
    'muteHttpExceptions': true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();

    if (responseCode === 200) {
      var responseBody = response.getContentText();
      var responseData = JSON.parse(responseBody);
      Logger.log('Password Data: ' + JSON.stringify(responseData, null, 2));
      // Retourneer data-object of null
      return responseData.data || null;
    } else {
      Logger.log('Error fetching password: ' + response.getContentText());
      return null;
    }
  } catch (error) {
    Logger.log('Error in getPasswordById: ' + error);
    return null;
  }
}

// ----------------------------
// API-call: Wachtwoord aanmaken
// ----------------------------

// From: PasswordAPICalls.js
function createPassword(organizationId, name, username, passwordCategoryId, password, map, notes, passurl) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var apiKey = scriptProperties.getProperty('IT_GLUE_API_KEY');

  if (!apiKey) {
    Logger.log('API Key is missing');
    Browser.msgBox('API Key is missing');
    return null;
  }

  var url = `https://api.eu.itglue.com/organizations/${organizationId}/relationships/passwords`;
  Logger.log('Create Password Request URL: ' + url);

  // ITGlue JSON: type + attributes
  var payload = {
    "data": {
      "type": "passwords",
      "attributes": {
        "organization-id": organizationId,
        "name": name,
        "username": username,
        "password-category-id": passwordCategoryId,
        "password": password.toString(),        // forceer string
        "password-folder-id": map,              // kan null zijn voor root
        "url": passurl,
        "notes": notes
      }
    }
  };

  var options = {
    'method': 'POST',
    'headers': {
      'Content-Type': 'application/vnd.api+json',
      'x-api-key': apiKey
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    Logger.log('Response Code for password creation: ' + responseCode);

    if (responseCode === 201) { // Created
      var responseData = JSON.parse(response.getContentText());
      return responseData.data; // Geef aangemaakte wachtwoorddata terug
    } else {
      Logger.log('API returned an error: ' + JSON.stringify(payload));
      Browser.msgBox('API returned an error: ' + response.getContentText());
      return null;
    }

  } catch (error) {
    Logger.log('Error creating password: ' + error);
    Browser.msgBox('Error creating password: ' + error);
    return null;
  }
}

// ----------------------------
// API-call: Map aanmaken
// ----------------------------

// From: PasswordAPICalls.js
function createPasswordFolder(organizationId, folderName, mapparent_id) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var sessionkey = scriptProperties.getProperty('IT_GLUE_SESSION_TOKEN');

  // Endpoint voor mappen-relationship van organisatie
  var url = `https://api.eu.itglue.com/organizations/${organizationId}/relationships/password_folders`;

  var payload = {
    "data": {
      "type": "password_folders",
      "attributes": {
        "organization-id": organizationId,
        "name": folderName,
        "parent-id": mapparent_id // null of ID
      }
    }
  };

  var options = {
    'method': 'POST',
    'headers': {
      'Content-Type': 'application/vnd.api+json',
      'Authorization': sessionkey // gebruikt session token
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    Logger.log('Response Code for folder creation: ' + responseCode);

    if (responseCode === 201) { // Created
      var data = JSON.parse(response.getContentText());
      return data.data.id; // Return het ID van de nieuwe map
    } else {
      Logger.log('API returned an error during folder creation: ' + response.getContentText());
    }
  } catch (error) {
    Logger.log('Error creating password folder: ' + error);
  }

  return null; // Geen map aangemaakt
}

// ----------------------------
// API-call: Map zoeken (via lijst + filteren op parent)
// ----------------------------

// From: PasswordAPICalls.js
function findPasswordFolder(organizationId, folderName, parentId) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var sessionKey = scriptProperties.getProperty('IT_GLUE_SESSION_TOKEN');

  if (!sessionKey) {
    Logger.log('API Key is missing');
    return [];
  }

  Logger.log('Organization ID:' + organizationId);
  Logger.log('folderName:' + folderName);
  Logger.log('parentId:' + parentId);

  // Ophalen van ALLE mappen van de organisatie; vervolgens client-side filteren
  var url = `https://api.eu.itglue.com/organizations/${organizationId}/relationships/password_folders`;

  var options = {
    'method': 'GET',
    'headers': {
      'Content-Type': 'application/vnd.api+json',
      'Authorization': sessionKey
    },
    'muteHttpExceptions': true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    if (responseCode === 200) {
      try {
        var responseData = JSON.parse(responseBody);

        if (responseData) {
          var folders = responseData.data;

          // Zoek exacte naam + dezelfde parent-id (let op: parentId kan null zijn)
          for (var i = 0; i < folders.length; i++) {
            var folder = folders[i];
            if (folder.attributes['name'] === folderName &&
              String(folder.attributes['parent-id']) === String(parentId)) {
              Logger.log('Folder bestaat al: ' + folder.id);
              return folder.id; // Found
            }
          }

          Logger.log('Folder bestaat nog niet');
          return null; // Niet gevonden
        } else {
          Logger.log('Geen geldige data gevonden in de response');
        }
      } catch (error) {
        Logger.log('Fout bij het parsen van de JSON-response: ' + error);
      }
    } else {
      Logger.log('Error fetching folders: ' + responseBody);
    }
  } catch (error) {
    Logger.log('Error in findPasswordFolder: ' + error);
  }

  return null;
}

// ----------------------------
// Recursief zoeken (indien hiërarchie in response aanwezig is)
// ----------------------------

// From: PasswordAPICalls.js
function searchFolderRecursive(folders, folderName, parentId) {
  for (var i = 0; i < folders.length; i++) {
    var folder = folders[i];
    var attributes = folder['attributes'];

    // Naam-match + (optioneel) parent-match
    if (attributes['name'] === folderName) {
      if (parentId === null || String(attributes['parent-id']) === String(parentId)) {
        return folder['id'];
      }
    }

    // Recursief door children
    if (folder.relationships && folder.relationships.children) {
      var childFolders = folder.relationships.children.data;
      var found = searchFolderRecursive(childFolders, folderName, folder.id);
      if (found) return found;
    }
  }
  return null;
}

// ----------------------------
// Alternatieve map-lookup met filters + UI-keuze bij meerdere matches
// ----------------------------

// From: PasswordAPICalls.js
function findPasswordFolderWithParentCheck(organizationId, folderName, parentFolderId) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var sessionKey = scriptProperties.getProperty('IT_GLUE_SESSION_TOKEN');

  // Server-side filters op naam (en optioneel parent-id)
  var url = `https://api.eu.itglue.com/organizations/${organizationId}/relationships/password_folders?filter[name]=${encodeURIComponent(folderName)}`;
  if (parentFolderId) {
    url += `&filter[parent-id]=${parentFolderId}`;
  }

  var options = {
    'method': 'GET',
    'headers': {
      'Content-Type': 'application/vnd.api+json',
      'Authorization': sessionKey
    },
    'muteHttpExceptions': true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    Logger.log('Response Code for folder search: ' + responseCode);

    if (responseCode === 200) {
      var data = JSON.parse(response.getContentText());
      Logger.log("Mapzoekresultaten: " + JSON.stringify(data));

      if (data.data && data.data.length > 0) {
        var matchingFolders = [];

        // Verzamel alle exacte matches op naam + (optioneel) parent-id
        for (var i = 0; i < data.data.length; i++) {
          var folder = data.data[i];
          var apiFolderName = folder.attributes.name.trim();
          var apiParentId = folder.relationships.parent.data ? folder.relationships.parent.data.id : null;

          if (apiFolderName.toLowerCase() === folderName.trim().toLowerCase() &&
            (!parentFolderId || apiParentId === parentFolderId)) {
            matchingFolders.push(folder);
          }
        }

        // Geen matches
        if (matchingFolders.length === 0) {
          Logger.log(`Geen map gevonden met naam "${folderName}" en parent-id "${parentFolderId}"`);
          return null;
        }

        // Eén match → direct gebruiken
        if (matchingFolders.length === 1) {
          Logger.log(`Eén match gevonden voor "${folderName}": ID ${matchingFolders[0].id}`);
          return matchingFolders[0].id;
        }

        // Meerdere matches → vraag gebruiker om keuze
        var ui = SpreadsheetApp.getUi();
        var choices = matchingFolders.map((folder, index) =>
          `${index + 1}: ${folder.attributes.name} (Parent ID: ${folder.relationships.parent.data ? folder.relationships.parent.data.id : 'Geen'})`
        ).join('\n');

        var response = ui.alert(
          'Meerdere mappen gevonden',
          `Er zijn meerdere mappen gevonden met de naam "${folderName}".\n\n${choices}\n\nVoer het nummer in van de gewenste map.`,
          ui.ButtonSet.OK_CANCEL
        );

        if (response === ui.Button.CANCEL) {
          throw new Error("Gebruiker heeft actie geannuleerd.");
        }

        // Prompt voor index
        var userInput = ui.prompt(
          "Kies een map",
          `Voer het nummer in van de gewenste map:\n\n${choices}`,
          ui.ButtonSet.OK
        ).getResponseText();

        var selectedIndex = parseInt(userInput, 10) - 1;

        // Ongeldige invoer → fout
        if (isNaN(selectedIndex) || selectedIndex < 0 || selectedIndex >= matchingFolders.length) {
          throw new Error("Ongeldige keuze. Probeer het opnieuw.");
        }

        Logger.log(`Geselecteerde map: ${matchingFolders[selectedIndex].attributes.name} (ID: ${matchingFolders[selectedIndex].id})`);
        return matchingFolders[selectedIndex].id;
      }
    } else {
      Logger.log('API returned an error during folder search: ' + response.getContentText());
    }
  } catch (error) {
    Logger.log('Error fetching password folders: ' + error);
  }

  return null; // Geen map gevonden
}

// ----------------------------
// Wachtwoorden uit API naar sheet schrijven
// ----------------------------

// From: Passwords.js
function writePasswordsToSheet(passwords, startFunctie, sheetname) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('startFunctie is: ' + " " + startFunctie)

  // Default naar 'test' (wordt overschreven als startFunctie 1 of 2 is)
  var passwordsheet = sheet.getSheetByName('test');

  // startFunctie '1' → werk met opgegeven sheetname (maak indien niet bestaat)
  if (startFunctie == '1') {
    passwordsheet = sheet.getSheetByName(sheetname);
    if (!passwordsheet) {
      passwordsheet = sheet.insertSheet(sheetname);
    }
  }

  // startFunctie '2' → idem
  if (startFunctie == '2') {
    passwordsheet = sheet.getSheetByName(sheetname);
    Logger.log('test 2');
    if (!passwordsheet) {
      passwordsheet = sheet.insertSheet(sheetname);
    }
  }

  // Cache voor organisatie-namen op ID (minder API-calls)
  var organizationCache = {};

  // Maak sheet leeg vóór schrijven
  passwordsheet.clear();

  // Geen data → stop
  if (passwords.length === 0) {
    Logger.log("Geen wachtwoorden om naar de sheet te schrijven.");
    return;
  }

  // Kolom-headers
  var headers = [
    "organization-name", "organization-id", "name", "username",
    "password-category-id", "password", "password-folder-id",
    "url", "notes", "password-id", "Full Path", "MapNiveau1", "MapNiveau2",
    "MapNiveau3", "MapNiveau4", "MapNiveau5", "MapNiveau6",
  ];
  passwordsheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Waarden 2D-array bouwen per password-record
  var values = passwords.map(function (password) {
    var organizationId = password.attributes["organization-id"];
    var passwordId = password['id'];
    var organizationName = getCachedOrganizationNameById(organizationId, organizationCache);

    // Extra call om plaintext/attributen op te halen
    var passwordtekstObj = getPasswordById(String(organizationId), String(passwordId));
    var folderId = password.attributes["password-folder-id"];

    return [
      organizationName || "",
      organizationId || "",
      password.attributes.name || "",
      password.attributes.username || "",
      password.attributes["password-category-id"] || "",
      // let op: passwordtekstObj kan null zijn
      passwordtekstObj && passwordtekstObj.attributes ? (passwordtekstObj.attributes.password || "") : "",
      folderId || "",
      password.attributes.url || "",
      password.attributes.notes || "",
      password.id || "",
      "", "", "", "", "", "", "" // Full Path + MapNiveaus worden later gevuld/gelaten
    ];
  });

  // Schrijf alle rijen in één keer
  passwordsheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

// ----------------------------
// Export: vraag organisatie, haal wachtwoorden, schrijf naar sheet
// ----------------------------

// From: Passwords.js
function exportPasswordsForOrganization(organisatieNaam, startFunctie, sheetname) {
  var ui = SpreadsheetApp.getUi(); // UI referentie

  var organizationName;

  // Als geen naam is meegegeven, vraag gebruiker om de naam
  if (!organisatieNaam) {
    var response = ui.prompt(
      "Organisatie naam",
      "Wat is de naam van de organisatie waarvoor je wachtwoorden wilt ophalen?",
      ui.ButtonSet.OK
    );
    organizationName = response.getResponseText();
  } else {
    // Bevestiging met voorgestelde naam
    var acceptlocation = ui.prompt(
      "Automatisch locatie ophalen",
      "Is de organisatie die u wilt ophalen uit IT-Glue: " + organisatieNaam + "?",
      ui.ButtonSet.OK_CANCEL
    );

    if (acceptlocation.getSelectedButton() == ui.Button.OK) {
      organizationName = organisatieNaam;
    } else {
      // Vraag opnieuw om naam
      var response2 = ui.prompt(
        "Organisatie naam",
        "Wat is de naam van de organisatie waarvoor je wachtwoorden wilt ophalen?",
        ui.ButtonSet.OK
      );
      organizationName = response2.getResponseText();
    }
  }

  // ID opzoeken via naam
  var organizationId = getOrganizationID(organizationName);

  // Niet gevonden → vraag retry of annuleren
  if (!organizationId) {
    var retryResponse = ui.alert(
      "Klant niet gevonden",
      "De organisatie '" + organizationName + "' kon niet worden gevonden. Probeert u het opnieuw?",
      ui.ButtonSet.OK_CANCEL
    );

    if (retryResponse == ui.Button.OK) {
      // Herstart zonder voorgedefinieerde naam
      exportPasswordsForOrganization(null, startFunctie);
    } else {
      Logger.log("Gebruiker heeft geannuleerd. Het proces is gestopt.");
      return;
    }
  }

  // Wachtwoorden ophalen
  var passwords = getPasswordsFromOrganization(organizationId);
  if (!passwords || passwords.length === 0) {
    Logger.log("Geen wachtwoorden gevonden voor organisatie: " + organizationName);
    return;
  }

  // Naar sheet schrijven (sheetname wordt doorgegeven)
  writePasswordsToSheet(passwords, startFunctie, sheetname);
  Logger.log("Wachtwoorden succesvol geëxporteerd naar de sheet.");
}

// ----------------------------
// Cache helper: org-naam per ID cachen
// ----------------------------

// From: Passwords.js
function getCachedOrganizationNameById(organizationId, organizationCache) {
  if (organizationCache[organizationId]) {
    return organizationCache[organizationId];
  } else {
    var organizationName = getOrganizationNameById(organizationId);
    organizationCache[organizationId] = organizationName;
    return organizationName;
  }
}

// ----------------------------
// Process: input → import + https toevoegen + validatie/kleuren
// ----------------------------

// From: Process password.js
function processPasswords(person) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Kies input- en import-tabbladen op basis van persoon
  var s1Sheet, s2Sheet, sheetname;
  if (person === "1") {
    s1Sheet = spreadsheet.getSheetByName("Input(Corné)");
    s2Sheet = spreadsheet.getSheetByName("Import(Corné)");
    sheetname = "Import(Corné)";
  } else if (person === "2") {
    s1Sheet = spreadsheet.getSheetByName("Input(Kevin)");
    s2Sheet = spreadsheet.getSheetByName("Import(Kevin)");
    sheetname = "Import(Kevin)";
  }

  // Doelblad leegmaken voor schone start
  clearSpecificSheet(sheetname);

  // Gewenste headers
  var headers = [
    "organization-name", "organization-id", "name", "username",
    "password-category-id", "password", "password-folder-id", "url", "notes", "password-id",
    "Full Path", "MapNiveau1", "MapNiveau2", "MapNiveau3", "MapNiveau4", "MapNiveau5", "MapNiveau6"
  ];

  // Controleer of headers al aanwezig zijn; zo niet → setten
  var existingHeaders = s2Sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var needToAddHeaders = !existingHeaders.some(header => header === headers[0]);
  if (needToAddHeaders) {
    s2Sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  // Data uit inputblad
  var data = s1Sheet.getDataRange().getValues();

  // Verwerk rijen (sla header over)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = row[0];
    var login = row[1];
    var password = row[2];
    var url = row[3];
    var note = row[4];
    var path = row[6];

    // Bouw nieuwe rij voor importblad (indices komen overeen met headers)
    var newRow = [];
    newRow[2] = name;
    newRow[3] = login;

    // Als password-formule voorkomt (=...), prefix met ' om formule te ontkrachten
    if (password && password.toString().startsWith("=")) {
      password = "'" + password.toString();
    }

    // Lege wachtwoorden naar placeholder
    if (password == "") {
      password = "--";
    }

    newRow[5] = password;
    newRow[7] = "'" + url; // prefix ' zodat Sheets de URL tekstueel behandelt
    newRow[8] = note;

    if (path) {
      // Splits pad op '#' (oude kluisnaam + rest)
      var split1 = path.split('#');

      if (split1.length > 1) {
        var remainder = split1[1]; // deel na '#'

        // Splits remainder op '/' → [kluis, klantnaam, map1, map2, ...]
        var split2 = remainder.split('/');

        // Klantnaam = eerste element NA de kluisnaam
        var klantnaam = split2[1];
        newRow[0] = klantnaam;

        // Mappen = de rest na klantnaam
        var folders = split2.slice(2);

        // Schrijf mapniveaus in kolommen MapNiveau1..6 en Full Path
        if (folders.length > 0) {
          for (var k = 0; k < folders.length; k++) {
            newRow[11 + k] = folders[k]; // MapNiveau1 begint op index 11
          }
          if (folders.length > 0) {
            newRow[10] = folders.join("/"); // Full Path
          }
        }
      }

      // Voeg rij toe aan importblad
      s2Sheet.appendRow(newRow);

      // Repareer URL's zonder protocol (https toevoegen waar nodig)
      voegHttpsToeAanURLs(sheetname);
    }
  }
}

// ----------------------------
// URL's repareren (https:// toevoegen) + daarna validatie/kleuren
// ----------------------------

// From: Process password.js
function voegHttpsToeAanURLs(sheetname) {
  Logger.log(sheetname);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetname);

  if (!sheet) {
    throw new Error("Tabblad niet gevonden: " + sheetname);
  }

  // Volledige data ophalen
  var data = sheet.getDataRange().getValues();

  // Zoek index van "url"-kolom
  var headerRow = data[0];
  var urlColumnIndex = headerRow.indexOf("url");

  if (urlColumnIndex === -1) {
    Logger.log("Kolom 'url' niet gevonden.");
    return;
  }

  // Regex voor IP-adres (optioneel poort) en domein zonder protocol
  var ipRegex = /^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)(:\d{1,5})?$/;
  var domainRegex = /^[a-zA-Z0-9-]+(\.[a-zA-Z0-9-]+)*\.[a-zA-Z]{2,}(:\d{1,5})?(\/.*)?$/;

  // Rijen doorlopen (vanaf 1 om header te skippen)
  for (var i = 1; i < data.length; i++) {
    var cellValue = data[i][urlColumnIndex];

    // Alleen behandelen als geen http/https is opgegeven
    if (cellValue && !cellValue.startsWith("http://") && !cellValue.startsWith("https://")) {
      if (ipRegex.test(cellValue) || domainRegex.test(cellValue)) {
        data[i][urlColumnIndex] = "https://" + cellValue;
      }
    }
  }

  // Data terugschrijven in één batch
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  // Na reparatie: valideer en kleur cellen
  valideerEnKleurCellen(sheetname);
}

// ----------------------------
// Validatie + cellen kleuren (verplicht, zwakke velden, duplicates, SSID, URLs)
// ----------------------------

// From: Process password.js
function valideerEnKleurCellen(sheetname) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetname);
  var data = sheet.getDataRange().getValues();
  var range = sheet.getDataRange();
  var backgroundColors = range.getBackgrounds(); // behoud bestaande kleuren

  // Benodigde kolommen
  var headers = data[0];
  var requiredColumns = ["organization-name", "name", "username", "password", "url", "Full Path"];
  var columnIndices = {};

  // Map kolomnamen → index
  for (var i = 0; i < headers.length; i++) {
    if (requiredColumns.includes(headers[i])) {
      columnIndices[headers[i]] = i;
    }
  }

  // Controle op aanwezigheid van verplichte kolommen
  requiredColumns.forEach(function (column) {
    if (columnIndices[column] === undefined) {
      throw new Error("Kolom '" + column + "' niet gevonden.");
    }
  });

  var nameIndex = columnIndices["name"];
  var fullPathIndex = columnIndices["Full Path"];
  var keyCounts = {};

  // Rijen valideren en kleuren (sla header over)
  for (var rowIndex = 1; rowIndex < data.length; rowIndex++) {
    var row = data[rowIndex];

    // Verplichte kolommen: leeg → oranje
    ["organization-name", "name", "username", "password"].forEach(function (column) {
      var colIndex = columnIndices[column];
      if (!row[colIndex] || row[colIndex].toString().trim() === "") {
        backgroundColors[rowIndex][colIndex] = "#FFA500"; // Oranje
      }
    });

    // Name minimaal 2 tekens, anders geel
    var nameValue = row[nameIndex];
    if (nameValue && nameValue.toString().trim().length < 2) {
      backgroundColors[rowIndex][nameIndex] = "#FFFF00"; // Geel
    }

    // Unieke sleutel van (name + Full Path) voor duplicate-detectie
    var fullPathValue = row[fullPathIndex] || "";
    var uniqueKey = `${nameValue}-${fullPathValue}`.trim();

    if (uniqueKey) {
      keyCounts[uniqueKey] = (keyCounts[uniqueKey] || []).concat(rowIndex);
    }

    // URL-check: http → rood, anders geen https → oranje
    var urlIndex = columnIndices["url"];
    var url = row[urlIndex];
    if (url) {
      url = url.toString().trim();
      if (url.startsWith("http://")) {
        backgroundColors[rowIndex][urlIndex] = "#FF0000"; // Rood
      } else if (!url.startsWith("https://")) {
        backgroundColors[rowIndex][urlIndex] = "#FFA500"; // Oranje
      }
    }

    // Naam bevat "SSID" (case-insensitive) → hele rij zachtgeel highlighten
    var nameValueLower = (row[nameIndex] || "").toString().toLowerCase();
    if (nameValueLower.includes("ssid")) {
      for (var col = 0; col < backgroundColors[rowIndex].length; col++) {
        backgroundColors[rowIndex][col] = "#FFFF99"; // Zachtgeel
      }
    }
  }

  // Markeer duplicaten (name + Full Path): lichtroze op name + Full Path
  Object.keys(keyCounts).forEach(function (key) {
    if (keyCounts[key].length > 1) {
      keyCounts[key].forEach(function (rowIndex) {
        backgroundColors[rowIndex][nameIndex] = "#FFC0CB"; // Licht roze
        backgroundColors[rowIndex][fullPathIndex] = "#FFC0CB";
      });
    }
  });

  // Kleuren terugschrijven
  range.setBackgrounds(backgroundColors);
}

// ----------------------------
// Sheet hulpfunctie: hele tabblad leegmaken
// ----------------------------

// From: Process password.js
function clearSpecificSheet(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  sheet.clear();
}
