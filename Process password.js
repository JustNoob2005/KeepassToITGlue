function processPasswords(person) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (person === "1") {
    var s1Sheet = spreadsheet.getSheetByName("Input(Corné)"); // Get the spreadsheet named "Second Migration"
    var s2Sheet = spreadsheet.getSheetByName("Import(Corné)"); // Get the spreadsheet named "Create passwordsheet-2"
    var sheetname = "Import(Corné)";
  }
  else if (person === "2") {
    var s1Sheet = spreadsheet.getSheetByName("Input(Kevin)"); // Get the spreadsheet named "Second Migration"
    var s2Sheet = spreadsheet.getSheetByName("Import(Kevin)"); // Get the spreadsheet named "Create passwordsheet-2"
    var sheetname = "Import(Kevin)";

  }


  clearSpecificSheet(sheetname); // To prevent any unwanted information in the target sheet , clear this one from all the information


  var headers = [ // Check if all headers are available
    "organization-name", "organization-id", "name", "username",
    "password-category-id", "password", "password-folder-id", "url", "notes", "password-id", "Full Path", "MapNiveau1", "MapNiveau2", "MapNiveau3", "MapNiveau4", "MapNiveau5", "MapNiveau6"
  ];

  var existingHeaders = s2Sheet.getRange(1, 1, 1, headers.length).getValues()[0]; // Get all the existingheaders.

  var needToAddHeaders = !existingHeaders.some(header => header === headers[0]); // Check if the headers need to be added
  if (needToAddHeaders) {
    s2Sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  var data = s1Sheet.getDataRange().getValues(); // Collect data from sheet s1.

  // For us there is a need to process only the data after the first split and onwards.
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = row[0];
    var login = row[1];
    var password = row[2];
    var url = row[3];
    var note = row[4];
    var path = row[6];

    var newRow = [];
    newRow[2] = name
    newRow[3] = login; // Kolom E

    if (password && password.toString().startsWith("=")) {
      password = "'" + password.toString(); // als string verwerken

    }
    newRow[5] = password; // Kolom G
    newRow[7] = "'" + url; // Kolom I, met enkelquote voor URL
    newRow[8] = note; // Kolom J

    if (path) { // If any path is specified this is used to split it into multiple folders.

      var split1 = path.split('#'); // Split on '#' sign

      if (split1.length > 1) {
        var remainder = split1[1]; // Only use the data after '#'

        var split2 = remainder.split('/'); //Split the remaining information on each /


        var klantnaam = split2[0]; // Set customer name to Colum B, Use the first item of the split2 variable as the value.
        newRow[0] = klantnaam;

        var folders = split2.slice(1); // This is used to ignore the first value in Split2, this is due the fact this was the name of the vault from our old password manager. , otherwise this creates a map on the top lvl of the customer.



        if (folders.length > 0) { // Check if there are multiple folders.
          for (var k = 0; k < folders.length; k++) { // Write away the values in Asc order of depth for each of the writeable levels. Do this in ascending order
            newRow[11 + k] = folders[k]; // Zet de waarde van folders[k] in de nieuwe rij op de juiste index
          }
          if (folders.length > 0) {
            newRow[10] = folders.join("/"); // Verbind de items met "/"
          }
        }


      }
      s2Sheet.appendRow(newRow); // Now add this data to the s2 Sheet as new row.
      voegHttpsToeAanURLs(sheetname); // Since our old data used some adresses or IP's without Https:// ITGlue did not recognize this as a valid URL link.

    }
  }
}

function voegHttpsToeAanURLs(sheetname) {
  Logger.log(sheetname);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetname); // Pas de sheetnaam indien nodig aan

  if (!sheet) {
    throw new Error("Tabblad niet gevonden: " + sheetname);
  }

  var values = sheet.getDataRange().getValues();
  // korte info

  var data = sheet.getDataRange().getValues();

  // Zoek de kolomindex voor de header "url"
  var headerRow = data[0];
  var urlColumnIndex = headerRow.indexOf("url");

  if (urlColumnIndex === -1) {
    Logger.log("Kolom 'url' niet gevonden.");
    return;
  }

  // Regex voor IP-adressen (met optioneel poortnummer) en domeinnamen zonder protocol
  var ipRegex = /^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)(:\d{1,5})?$/;
  var domainRegex = /^[a-zA-Z0-9-]+(\.[a-zA-Z0-9-]+)*\.[a-zA-Z]{2,}(:\d{1,5})?(\/.*)?$/;

  // Loop door de rijen (start bij 1 om de header over te slaan)
  for (var i = 1; i < data.length; i++) {
    var cellValue = data[i][urlColumnIndex];

    if (cellValue && !cellValue.startsWith("http://") && !cellValue.startsWith("https://")) {
      if (ipRegex.test(cellValue) || domainRegex.test(cellValue)) {
        // Voeg 'https://' toe als het een IP-adres of domeinnaam zonder protocol is
        data[i][urlColumnIndex] = "https://" + cellValue;
      }
    }
  }

  // Schrijf de gewijzigde data terug naar de sheet
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  valideerEnKleurCellen(sheetname);
}


function valideerEnKleurCellen(sheetname) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetname); // Pas de sheetnaam indien nodig aan
  var data = sheet.getDataRange().getValues();
  var range = sheet.getDataRange();
  var backgroundColors = range.getBackgrounds(); // Haal bestaande achtergrondkleuren op

  // Zoek de kolomindices van de headers
  var headers = data[0];
  var requiredColumns = ["organization-name", "name", "username", "password", "url", "Full Path"];
  var columnIndices = {};

  for (var i = 0; i < headers.length; i++) {
    if (requiredColumns.includes(headers[i])) {
      columnIndices[headers[i]] = i;
    }
  }

  // Controleer of alle benodigde kolommen bestaan
  requiredColumns.forEach(function (column) {
    if (columnIndices[column] === undefined) {
      throw new Error("Kolom '" + column + "' niet gevonden.");
    }
  });

  var nameIndex = columnIndices["name"];
  var fullPathIndex = columnIndices["Full Path"];

  var keyCounts = {};

  // Controleer de rijen (sla de header over)
  for (var rowIndex = 1; rowIndex < data.length; rowIndex++) {
    var row = data[rowIndex];

    // Controleer verplichte kolommen op lege waarden
    ["organization-name", "name", "username", "password"].forEach(function (column) {
      var colIndex = columnIndices[column];
      if (!row[colIndex] || row[colIndex].toString().trim() === "") {
        backgroundColors[rowIndex][colIndex] = "#FFA500"; // Oranje
      }
    });

    // Controleer of "name" minstens 2 tekens bevat
    var nameValue = row[nameIndex];
    if (nameValue && nameValue.toString().trim().length < 2) {
      backgroundColors[rowIndex][nameIndex] = "#FFFF00"; // Geel
    }

    // Combineer name en Full Path om een unieke sleutel te maken
    var fullPathValue = row[fullPathIndex] || ""; // Zorg ervoor dat lege waarden worden behandeld
    var uniqueKey = `${nameValue}-${fullPathValue}`.trim();

    // Tel het aantal keren dat de unieke sleutel voorkomt
    if (uniqueKey) {
      keyCounts[uniqueKey] = (keyCounts[uniqueKey] || []).concat(rowIndex);
    }

    // Controleer de URL
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
    // Highlight volledige rij als "SSID" voorkomt in de naam (hoofdletterongevoelig)
    var nameValueLower = (row[nameIndex] || "").toString().toLowerCase();
    if (nameValueLower.includes("ssid")) {
      for (var col = 0; col < backgroundColors[rowIndex].length; col++) {
        backgroundColors[rowIndex][col] = "#FFFF99"; // Zachtgeel
      }
    }
  }

  // Markeer dubbele rijen op basis van de unieke sleutel
  Object.keys(keyCounts).forEach(function (key) {
    if (keyCounts[key].length > 1) {
      keyCounts[key].forEach(function (rowIndex) {
        backgroundColors[rowIndex][nameIndex] = "#FFC0CB"; // Licht rood
        backgroundColors[rowIndex][fullPathIndex] = "#FFC0CB"; // Markeer ook Full Path
      });
    }
  });

  // Stel de nieuwe achtergrondkleuren in
  range.setBackgrounds(backgroundColors);
}


function clearSpecificSheet(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName); // Vervang 'Sheet1' door de naam van je sheet
  sheet.clear(); // Alles leegmaken
}