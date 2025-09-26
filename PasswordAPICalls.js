function getPasswordsFromOrganization(organizationId) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var apiKey = scriptProperties.getProperty('IT_GLUE_API_KEY');

  if (!apiKey) {
    Logger.log('API Key is missing');
    return [];
  }

  // Correct URL voor passwords op basis van organisatie ID
  var url = `https://api.eu.itglue.com/passwords?filter[organization_id]=${organizationId}&page[number]=1`; // Startpagina is 1
  var allPasswords = [];
  
  try {
    var hasMorePages = true;
    var currentPage = 1;

    while (hasMorePages) {
      var options = {
        'method': 'GET',
        'headers': {
          'Content-Type': 'application/vnd.api+json',
          'x-api-key': apiKey
        },
        'muteHttpExceptions': true
      };

      // Pas de URL aan voor de huidige pagina
      var pageUrl = `https://api.eu.itglue.com/passwords?filter[organization_id]=${organizationId}&page[number]=${currentPage}`;
    //  Logger.log('Request URL: ' + pageUrl);

      var response = UrlFetchApp.fetch(pageUrl, options);
      var responseCode = response.getResponseCode();
     // Logger.log('Response Code: ' + responseCode);

      if (responseCode !== 200) {
     //   Logger.log('API returned an error: ' + response.getContentText());
        return [];
      }

      var responseText = response.getContentText();
      if (!responseText || responseText.trim() === '') {
    //    Logger.log('Response is empty');
        return [];
      }

      var data = JSON.parse(responseText);
      var passwords = data.data || [];

      // Voeg de opgehaalde wachtwoorden toe aan de lijst
      allPasswords = allPasswords.concat(passwords);

      // Controleer of er meer pagina's zijn
      hasMorePages = data.links && data.links.next;
      currentPage++;
    }

    return allPasswords;
    
  } catch (error) {
    Logger.log('Error fetching passwords from IT Glue API: ' + error);
    return [];
  }
}


function getPasswordById(organizationId, passwordId) {
var scriptProperties = PropertiesService.getScriptProperties();
 var apiKey = scriptProperties.getProperty('IT_GLUE_API_KEY');

  if (!apiKey) {
    Logger.log('API Key is missing');
    return [];
  }

  

  // De URL voor de GET-aanroep naar de API om een wachtwoord op te halen
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
      
      // Je kunt hier de specifieke gegevens teruggeven die je nodig hebt
      return responseData.data || null; // Return de wachtwoordgegevens
    } else {
      Logger.log('Error fetching password: ' + response.getContentText());
      return null;
    }
  } catch (error) {
    Logger.log('Error in getPasswordById: ' + error);
    return null;
  }
}


function createPassword(organizationId, name, username, passwordCategoryId,password,map,notes,passurl) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var apiKey = scriptProperties.getProperty('IT_GLUE_API_KEY');
 

  if (!apiKey) {
    Logger.log('API Key is missing');
    Browser.msgBox('API Key is missing');
    return null;
  }

  var url = `https://api.eu.itglue.com/organizations/${organizationId}/relationships/passwords`;
  Logger.log('Create Password Request URL: ' + url);

  var payload = {
    "data": {
      "type": "passwords",
      "attributes": {
        "organization-id": organizationId,
        "name": name,
        "username": username,
        "password-category-id": passwordCategoryId,
        "password": password.toString(),
        "password-folder-id": map,
        "url": passurl,
        "notes":notes,
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

    if (responseCode === 201) { // 201 Created
      var responseData = JSON.parse(response.getContentText());
      return responseData.data; // Teruggeven van de gegevens van het aangemaakte wachtwoord
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





function createPasswordFolder(organizationId, folderName,mapparent_id) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var sessionkey = scriptProperties.getProperty('IT_GLUE_SESSION_TOKEN');
  
 // Logger.log('Nog even een check vanuit de createpasswordfolder'+ mapparent_id);

  var url = `https://api.eu.itglue.com/organizations/${organizationId}/relationships/password_folders`;
  var payload = {
    "data": {
      "type": "password_folders",
      "attributes": {
        "organization-id": organizationId,
        "name": folderName,
        "parent-id": mapparent_id
      }
    }
  };
  
  var options = {
    'method': 'POST',
    'headers': {
      'Content-Type': 'application/vnd.api+json',
       'Authorization': sessionkey
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    Logger.log('Response Code for folder creation: ' + responseCode);

    if (responseCode === 201) { // 201 Created
      var data = JSON.parse(response.getContentText());
      return data.data.id; // Teruggeven van het ID van de aangemaakte map
    } else {
      Logger.log('API returned an error during folder creation: ' + response.getContentText());
    }
  } catch (error) {
    Logger.log('Error creating password folder: ' + error);
  }

  return null; // Geen map aangemaakt
}

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

  // De URL voor de GET-aanroep naar de API om folders op te halen
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
    //Logger.log('Response Code: ' + responseCode);

    // Log de volledige response voor debuggen
    var responseBody = response.getContentText();
    //Logger.log('Response Body: ' + responseBody);

    if (responseCode === 200) {
      try {
        var responseData = JSON.parse(responseBody);
       // Logger.log('Parsed Response Data: ' + JSON.stringify(responseData, null, 2));

        // Controleer de hele structuur van responseData
        if (responseData) {
          var folders = responseData.data;
        //  Logger.log('Folders: ' + JSON.stringify(folders));

          // Zoek naar de folder in de opgehaalde gegevens
          for (var i = 0; i < folders.length; i++) {
            var folder = folders[i];
          //  Logger.log('Folder ' + (i + 1) + ': ' + JSON.stringify(folder, null, 2));

            // Controleer op de folder naam en parent-id
            if (folder.attributes['name'] === folderName && String(folder.attributes['parent-id']) === String(parentId)) {
              Logger.log('Folder bestaat al: ' + folder.id);
              return folder.id; // Return het folder ID als het bestaat
            }
          }
          
          Logger.log('Folder bestaat nog niet');
          return null; // Als de folder niet bestaat, return null
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



function searchFolderRecursive(folders, folderName, parentId) {
  for (var i = 0; i < folders.length; i++) {
    var folder = folders[i];
    var attributes = folder['attributes'];
  
    // Check of de naam overeenkomt:
    if (attributes['name'] === folderName) {
       if (parentId === null || String(attributes['parent-id']) === String(parentId)) {
        return folder['id'];
      }
    }
    // Check subfolders
    if (folder.relationships && folder.relationships.children) {
      var childFolders = folder.relationships.children.data;
      var found = searchFolderRecursive(childFolders, folderName, folder.id);
      if (found) return found;
    }
  }
  return null;
}


function findPasswordFolderWithParentCheck(organizationId, folderName, parentFolderId) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var sessionKey = scriptProperties.getProperty('IT_GLUE_SESSION_TOKEN');
  
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

        // Verzamel alle matches
        for (var i = 0; i < data.data.length; i++) {
          var folder = data.data[i];
          var apiFolderName = folder.attributes.name.trim();
          var apiParentId = folder.relationships.parent.data ? folder.relationships.parent.data.id : null;

          if (apiFolderName.toLowerCase() === folderName.trim().toLowerCase() && (!parentFolderId || apiParentId === parentFolderId)) {
            matchingFolders.push(folder);
          }
        }

        // Als geen matches, return null
        if (matchingFolders.length === 0) {
          Logger.log(`Geen map gevonden met naam "${folderName}" en parent-id "${parentFolderId}"`);
          return null;
        }

        // Als één match, gebruik deze
        if (matchingFolders.length === 1) {
          Logger.log(`Eén match gevonden voor "${folderName}": ID ${matchingFolders[0].id}`);
          return matchingFolders[0].id;
        }

        // Meerdere matches, vraag gebruiker om keuze
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

        // Vraag gebruiker naar een nummer
        var userInput = ui.prompt(
          "Kies een map",
          `Voer het nummer in van de gewenste map:\n\n${choices}`,
          ui.ButtonSet.OK
        ).getResponseText();

        var selectedIndex = parseInt(userInput, 10) - 1;

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

