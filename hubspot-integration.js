/**
 * Updates student records in a Google Sheet with additional properties fetched from HubSpot.
 * 
 * For each student (identified by email), this function fetches values for the specified properties
 * from the HubSpot CRM and writes them into the corresponding columns of the sheet.
 *
 * @param {Array<string>} properties - List of HubSpot property names to retrieve and write into the sheet. 
 */
function updatePropertiesInSheet(properties) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("all-availabilities"); // Replace with your sheet name
  const data = sheet.getDataRange().getValues(); // Get all data from the sheet
  const headers = data[0]; // Assuming the first row contains headers
  const emailIndex = headers.indexOf("Email Address"); // Column with emails

  if (emailIndex === -1) {
    Logger.log("Please ensure 'Email' column is present.");
    return;
  }

  // Find column indexes for each property
  const propertyIndexes = properties.map(property => headers.indexOf(property));
  if (propertyIndexes.includes(-1)) {
    Logger.log("Please ensure all properties have corresponding columns in the sheet.");
    return;
  }

  for (let i = 1; i < data.length; i++) {
    const email = data[i][emailIndex]; // Get the email from the current row
    if (email) {
      const propertyValues = getPropertiesFromHubSpot(email, properties); // Get property values from HubSpot
      if (propertyValues) {
        // Insert values into the appropriate columns
        propertyValues.forEach((value, index) => {
          sheet.getRange(i + 1, propertyIndexes[index] + 1).setValue(value);
        });
      }
      Utilities.sleep(500); // 0.5-second delay to prevent rate limit errors
    }
  }
}

/**
 * Fetches specified property values for a contact from HubSpot based on email address.
 * 
 * Uses the HubSpot CRM v3 Search API to find the contact and retrieve values for the given properties.
 *
 * @param {string} email - Email address used to identify the contact in HubSpot.
 * @param {Array<string>} properties - List of property names to retrieve.
 * @returns {Array<string> | null} - Array of property values in the same order as requested, or null if contact not found.
 */
function getPropertiesFromHubSpot(email, properties) {
  const url = `https://api.hubapi.com/crm/v3/objects/contacts/search`; // HubSpot API endpoint for searching contacts
  const apiKey = 'REDACTED'; // Use your private app token
  const payload = {
    filterGroups: [{
      filters: [{
        propertyName: "email",
        operator: "EQ",
        value: email
      }]
    }],
    properties: properties // Pass the properties argument
  };

  const options = {
    method: 'POST',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + apiKey },
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (json && json.results && json.results.length > 0) {
    // Return an array of property values corresponding to the passed properties
    return properties.map(prop => json.results[0].properties[prop] || ""); // Default to empty string if no value
  } else {
    return null;
  }
}

/**
 * Example usage function to populate major and minor fields in the sheet.
 * 
 * Can be modified to call `updatePropertiesInSheet()` with any set of HubSpot properties
 * that are also represented as column headers in the target sheet.
 */
function integrateMain() {
  updatePropertiesInSheet(["primary_field_of_study___major", "secondary_field_of_study___minor"]);
}
