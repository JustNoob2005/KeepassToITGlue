function getOrganizationID(organizationName) {
  // Assume getOrganizationIdByName makes the API call and returns the ID if found
  
  var organizationId = getOrganizationIdByName(organizationName);
  if (!organizationId) {
    Logger.log(`Kan geen organisatie-ID vinden voor: ${organizationName}`);
    return null;
  } else {
    Logger.log("Organizations - 9 - Organisatie-ID voor " + organizationName + ": " + organizationId);
    return organizationId;
  }


}

function getOrganiations(){

  var organisations = getOrganizationOverview();

  Logger.log(JSON.stringify(organisations,null,2));
}