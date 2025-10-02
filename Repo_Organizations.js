// ----------------------------
// Organization repository: wraps ITGlue organization API calls
// ----------------------------

/** Auto-generated restructure. */

/**
 * ----------------------------
 * ORGANIZATION REPOSITORY
 * ----------------------------
 * 
 * Dit script vormt een "repository"-laag voor ITGlue-organisatiegegevens.
 * Het bevat functies die communiceren met de ITGlue API om organisaties
 * op te zoeken en informatie terug te geven (ID of naam).
 * 
 * Vereisten:
 * - Een geldige IT Glue API key moet in de Script Properties staan
 *   onder de naam: "IT_GLUE_API_KEY".
 *   (In Apps Script: Bestand > Projectinstellingen > Script properties).
 * 
 * Functies:
 * 
 * getOrganizationIdByName(organizationName)
 *   → Zoekt de organisatie-ID op basis van de organisatienaam.
 *   → Retourneert het ID of null als niet gevonden.
 * 
 * getOrganizationNameById(organizationId)
 *   → Zoekt de organisatienaam op basis van de organisatie-ID.
 *   → Retourneert de naam of null als niet gevonden.
 * 
 * getOrganizationID(organizationName)
 *   → Wrapper die getOrganizationIdByName aanroept en logging toevoegt.
 *   → Retourneert het ID of null.
 * 
 * getOrganiations()
 *   → Roept (extern gedefinieerde) getOrganizationOverview() aan
 *     en logt het resultaat.
 *   → Nuttig voor debugging en overzicht.
 * 
 * Gebruik:
 * - Deze functies worden meestal aangeroepen vanuit hogere logica
 *   (bijvoorbeeld bij het verwerken van klantgegevens of wachtwoorden).
 * - getOrganizationIdByName en getOrganizationNameById zijn de kern-API-calls.
 * - De wrappers (getOrganizationID, getOrganiations) dienen vooral
 *   voor logging of om data overzichtelijker te maken.
 * 
 * Foutafhandeling:
 * - Bij ontbrekende API key of API-fouten geven functies null terug.
 * - Logging in Logger.log helpt bij debugging (Apps Script console).
 */


// ----------------------------
// Zoek een organisatie-ID op basis van de naam
// ----------------------------
function getOrganizationIdByName(organizationName) {
  // Haal API key op uit de Script Properties (moet eerder ingesteld zijn)
  var scriptProperties = PropertiesService.getScriptProperties();
  var apiKey = scriptProperties.getProperty('IT_GLUE_API_KEY');

  Logger.log(organizationName);

  // Controleer of de API key aanwezig is
  if (!apiKey) {
    Logger.log('API Key is missing');
    return null;
  }

  // Endpoint voor het zoeken van een organisatie op naam
  var url = `https://api.eu.itglue.com/organizations?filter[name]=${encodeURIComponent(organizationName)}`;

  // HTTP request opties
  var options = {
    'method': 'GET',
    'headers': {
      'Content-Type': 'application/vnd.api+json',
      'x-api-key': apiKey
    },
    'muteHttpExceptions': true // Zorgt dat de functie niet crasht bij HTTP-fouten
  };

  try {
    // Voer API-call uit
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    Logger.log('Response Code for organization search: ' + responseCode);

    if (responseCode === 200) {
      // Succesvolle call → data parsen
      var data = JSON.parse(response.getContentText());

      if (data.data && data.data.length > 0) {
        // Retourneer de ID van de eerste gevonden organisatie
        return data.data[0].id;
      } else {
        // Geen organisatie gevonden met de opgegeven naam
        Logger.log(`Geen organisatie gevonden met de naam: ${organizationName}`);
        return null;
      }
    } else {
      // API gaf een fout terug (geen 200-status)
      Logger.log('API returned an error: ' + response.getContentText());
      return null;
    }
  } catch (error) {
    // Exception tijdens API-call (bijv. netwerkfout)
    Logger.log('Error fetching organization by name: ' + error);
    return null;
  }
}

// ----------------------------
// Zoek een organisatienaam op basis van de ID
// ----------------------------
function getOrganizationNameById(organizationId) {
  // Haal API key op uit de Script Properties
  var scriptProperties = PropertiesService.getScriptProperties();
  var apiKey = scriptProperties.getProperty('IT_GLUE_API_KEY');

  // Controleer of de API key aanwezig is
  if (!apiKey) {
    Logger.log('API Key is missing');
    return null;
  }

  // Endpoint om organisatiegegevens op te halen via ID
  var url = `https://api.eu.itglue.com/organizations/${organizationId}`;

  // HTTP request opties
  var options = {
    'method': 'GET',
    'headers': {
      'Content-Type': 'application/vnd.api+json',
      'x-api-key': apiKey
    },
    'muteHttpExceptions': true
  };

  try {
    // Voer API-call uit
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    Logger.log('Response Code for organization fetch: ' + responseCode);

    if (responseCode === 200) {
      // Succesvolle call → data parsen
      var data = JSON.parse(response.getContentText());

      // Controleer of er een naam in de response zit
      if (data.data && data.data.attributes && data.data.attributes['name']) {
        // Retourneer de organisatienaam
        Logger.log(JSON.stringify(data), null, 2);
        return data.data.attributes['name'];
      } else {
        Logger.log(`Geen organisatiegegevens gevonden voor ID: ${organizationId}`);
        return null;
      }
    } else {
      // API gaf een fout terug
      Logger.log('API returned an error: ' + response.getContentText());
      return null;
    }
  } catch (error) {
    // Exception tijdens API-call
    Logger.log('Error fetching organization by ID: ' + error);
    return null;
  }
}

// ----------------------------
// Hulpfunctie om organisatie-ID op te halen (wrapper)
// ----------------------------
function getOrganizationID(organizationName) {
  // Roep getOrganizationIdByName() aan om ID op te halen
  var organizationId = getOrganizationIdByName(organizationName);

  if (!organizationId) {
    // Geen ID gevonden
    Logger.log(`Kan geen organisatie-ID vinden voor: ${organizationName}`);
    return null;
  } else {
    // ID gevonden → log en return
    Logger.log("Organizations - 9 - Organisatie-ID voor " + organizationName + ": " + organizationId);
    return organizationId;
  }
}

// ----------------------------
// Haal een overzicht van organisaties op (wrapper)
// ----------------------------
function getOrganiations() {
  // Roept een externe functie aan (nog niet gedefinieerd hier): getOrganizationOverview()
  var organisations = getOrganizationOverview();

  // Log het complete overzicht als JSON
  Logger.log(JSON.stringify(organisations, null, 2));
}

/**
 * Fetch a full overview (all pages) of ITGlue organizations.
 * Returns: Array<{ id: string, name: string }>
 * Requires Script Property: IT_GLUE_API_KEY
 */
function getOrganizationOverview() {
  var apiKey = CONFIG_get('IT_GLUE_API_KEY');
  if (!apiKey) {
    Logger.log('Missing Script Property: IT_GLUE_API_KEY');
    return [];
  }

  var baseUrl = 'https://api.eu.itglue.com/organizations';
  var pageSize = 100; // ITGlue supports JSON:API-style pagination. 100 is a common practical max.

  var all = [];
  var page = 1;

  while (true) {
    var url = baseUrl + '?page[size]=' + pageSize + '&page[number]=' + page +
      '&fields[organizations]=name'; // only fetch what we need
    var options = {
      method: 'get',
      headers: {
        'Content-Type': 'application/vnd.api+json',
        'x-api-key': apiKey
      },
      muteHttpExceptions: true
    };

    var resp = UrlFetchApp.fetch(url, options);
    var code = resp.getResponseCode();
    if (code !== 200) {
      Logger.log('ITGlue organizations fetch failed (HTTP ' + code + '): ' + resp.getContentText());
      break;
    }

    var json = JSON.parse(resp.getContentText());
    var data = (json && json.data) || [];
    if (!data.length) {
      break; // no more rows
    }

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (row && row.id && row.attributes && typeof row.attributes.name === 'string') {
        all.push({ id: String(row.id), name: row.attributes.name });
      }
    }

    // Stop if we received less than a full page.
    if (data.length < pageSize) {
      break;
    }
    page++;
  }

  Logger.log('Fetched organizations: ' + all.length);
  return all;
}

