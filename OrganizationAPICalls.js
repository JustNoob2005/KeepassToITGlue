function getOrganizationIdByName(organizationName) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var apiKey = scriptProperties.getProperty('IT_GLUE_API_KEY');

  Logger.log(organizationName);

  if (!apiKey) {
    Logger.log('API Key is missing');
    return null;
  }

  // Endpoint voor het zoeken van een organisatie op naam
  var url = `https://api.eu.itglue.com/organizations?filter[name]=${encodeURIComponent(organizationName)}`;

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
    Logger.log('Response Code for organization search: ' + responseCode);

    if (responseCode === 200) {
      var data = JSON.parse(response.getContentText());
      if (data.data && data.data.length > 0) {
        // Retourneer de ID van de eerste gevonden organisatie
        return data.data[0].id;
      } else {
        Logger.log(`Geen organisatie gevonden met de naam: ${organizationName}`);
        return null;
      }
    } else {
      Logger.log('API returned an error: ' + response.getContentText());
      return null;
    }
  } catch (error) {
    Logger.log('Error fetching organization by name: ' + error);
    return null;
  }
}

function getOrganizationNameById(organizationId) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var apiKey = scriptProperties.getProperty('IT_GLUE_API_KEY');

  if (!apiKey) {
    Logger.log('API Key is missing');
    return null;
  }

  // Endpoint om een organisatie op te halen op basis van ID
  var url = `https://api.eu.itglue.com/organizations/${organizationId}`;

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
    Logger.log('Response Code for organization fetch: ' + responseCode);

    if (responseCode === 200) {
      var data = JSON.parse(response.getContentText());
      if (data.data && data.data.attributes && data.data.attributes['name']) {
        // Retourneer de naam van de organisatie
        Logger.log(JSON.stringify(data),null,2);
        return data.data.attributes['name'];
      } else {
        Logger.log(`Geen organisatiegegevens gevonden voor ID: ${organizationId}`);
        return null;
      }
    } else {
      Logger.log('API returned an error: ' + response.getContentText());
      return null;
    }
  } catch (error) {
    Logger.log('Error fetching organization by ID: ' + error);
    return null;
  }
}