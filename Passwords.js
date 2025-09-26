function writePasswordsToSheet(passwords,startFunctie, sheetname) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('startFunctie is: ' + " " + startFunctie)

  var passwordsheet = sheet.getSheetByName('test');

  if (startFunctie == '1'){
    var passwordsheet = sheet.getSheetByName(sheetname);
  if (!passwordsheet) {
    passwordsheet = sheet.insertSheet(sheetname);
  }
  }

  if (startFunctie == '2'){
  var passwordsheet = sheet.getSheetByName(sheetname);
   Logger.log('test 2');
  if (!passwordsheet) {
    passwordsheet = sheet.insertSheet(sheetname);
  }
  }


  var organizationCache = {};


  passwordsheet.clear(); // Sheet leegmaken voordat nieuwe data wordt toegevoegd

  if (passwords.length === 0) {
    Logger.log("Geen wachtwoorden om naar de sheet te schrijven.");
    return;
  }

  // Headers volgens jouw specificatie
  var headers = [
    "organization-name", "organization-id", "name", "username",
    "password-category-id", "password", "password-folder-id",
    "url", "notes", "password-id","Full Path", "MapNiveau1","MapNiveau2",
    "MapNiveau3","MapNiveau4","MapNiveau5","MapNiveau6",
  ];
  passwordsheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Waarden in een 2D-array zetten

//Logger.log(passwords[0],null,2);

  var values = passwords.map(function(password) {
    var organizationId = password.attributes["organization-id"];
    var passwordId = password['id'];
    var organizationName = getCachedOrganizationNameById(organizationId, organizationCache);

    var passwordtekst = getPasswordById(String(organizationId),String(passwordId))
    var folderId = password.attributes["password-folder-id"];

    return [
      organizationName || "",
      organizationId || "",
      password.attributes.name || "",
      password.attributes.username || "",
      password.attributes["password-category-id"] || "",
      passwordtekst.attributes.password || "",
      folderId || "",
      password.attributes.url || "",
      password.attributes.notes || "",
      password.id || "",
     "",
     "",
     "",
     "",
     "",
     "",
     "",
    ];
  });

  // Schrijf de gegevens naar de sheet
  passwordsheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

function exportPasswordsForOrganization(organisatieNaam, startFunctie,sheetname) {
  var ui = SpreadsheetApp.getUi(); // Verkrijg de UI van de Spreadsheet

  // Als organisatieNaam leeg is, vraag dan om de naam van de organisatie
  if (!organisatieNaam) {
    var response = ui.prompt(
      "Organisatie naam",
      "Wat is de naam van de organisatie waarvoor je wachtwoorden wilt ophalen?",
      ui.ButtonSet.OK
    );
    var organizationName = response.getResponseText(); // Verkrijg het antwoord
  } else {
    var acceptlocation = ui.prompt(
      "Automatisch locatie ophalen",
      "Is de organisatie die u wilt ophalen uit IT-Glue: " + organisatieNaam + "?",
      ui.ButtonSet.OK_CANCEL
    );
    
    // Controleer of de gebruiker op 'OK' heeft geklikt
    if (acceptlocation.getSelectedButton() == ui.Button.OK) {
      var organizationName = organisatieNaam;
    } else {
      var response = ui.prompt(
        "Organisatie naam",
        "Wat is de naam van de organisatie waarvoor je wachtwoorden wilt ophalen?",
        ui.ButtonSet.OK
      );
      var organizationName = response.getResponseText(); // Verkrijg het antwoord
    }
  }

  //Logger.log(organizationName); // Voor debugging

  // Verkrijg de organisatie ID op basis van de naam
  var organizationId = getOrganizationID(organizationName);
  
  // Als de organisatie niet gevonden is, toon dan een pop-up
  if (!organizationId) {
    var retryResponse = ui.alert(
      "Klant niet gevonden", 
      "De organisatie '" + organizationName + "' kon niet worden gevonden. Probeert u het opnieuw?", 
      ui.ButtonSet.OK_CANCEL
    );
    
    if (retryResponse == ui.Button.OK) {
      // Herhaal het proces door opnieuw om de organisatie naam te vragen
      exportPasswordsForOrganization(null,startFunctie);
    } else {
      Logger.log("Gebruiker heeft geannuleerd. Het proces is gestopt.");
      return;  // Stop het proces als de gebruiker op "Cancel" klikt
    }
  }

  // Verkrijg de wachtwoorden voor de organisatie als de organisatie wel gevonden is
  var passwords = getPasswordsFromOrganization(organizationId);
  if (!passwords || passwords.length === 0) {
    Logger.log("Geen wachtwoorden gevonden voor organisatie: " + organizationName);
    return;
  }


  // Schrijf de wachtwoorden naar de sheet en voer verdere processen uit
  writePasswordsToSheet(passwords, startFunctie);
  Logger.log("Wachtwoorden succesvol geëxporteerd naar de sheet.");
}

function getCachedOrganizationNameById(organizationId, organizationCache) {
  // Controleer of de organisatie al in de cache staat
  if (organizationCache[organizationId]) {
    // Als de organisatie al in de cache staat, geef de naam terug
    return organizationCache[organizationId];
  } else {
    // Anders, haal de naam op en voeg deze toe aan de cache
    var organizationName = getOrganizationNameById(organizationId); // Oorspronkelijke functie
    organizationCache[organizationId] = organizationName; // Voeg de naam toe aan de cache
    return organizationName;
  }
}