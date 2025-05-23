/**
 * Returns the top N most frequently occurring items in an array.
 *
 * @param {Array<string>} array - The input array of strings to analyze.
 * @param {number} topN - The number of top occurrences to return.
 * @returns {Array<string>} - An array of the top N most common items, sorted by frequency.
 */
function getTopOccurrences(array, topN) {
  let countMap = array.reduce((acc, item) => {
    acc[item] = (acc[item] || 0) + 1;
    return acc;
  }, {});
  return Object.entries(countMap)
    .sort((a, b) => b[1] - a[1])
    .slice(0, topN)
    .map(([key]) => key);
}

/**
 * Returns the full row of student data for a given email.
 *
 * @param {string} email - The student's email address.
 * @param {Array<Array<any>>} data - All student data as a 2D array from getValues().
 * @returns {Array<any> | null} - The student's data row, or null if email is not found.
 */
function getRowByEmail(email, data) {
  for (let i = 0; i < data.length; i++) {
    if (data[i][3] === email) {  // email is in column 4 (index 3)
      return data[i];
    }
  }
  return null;
}

/**
 * Given a student's email and full student data array, returns the student's first and last name.
 *
 * @param {string} email - The student's email address.
 * @param {Array<Array<any>>} data - All student data as a 2D array from getValues().
 * @returns {Array<string> | string} - First and last name, or "Email not found" if no match.
 */
function getNameByEmail(email, data) {
  // Assuming the first name is in the second column (index 1), last name in third column (index 2)
  // and the email is in the fourth column (index 3).
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][3] === email) {  // Compare the email in the fourth column
      return [data[i][1], data[i][2]];         // Return the first and last name in an array
    }
  }
  
  // If email is not found
  return "Email not found";
}

/**
 * Retrieves a comma-separated list of emails from a specific cell in a sheet,
 * and returns them as a cleaned array of trimmed strings.
 *
 * @param {string} sheetName - The name of the sheet to access.
 * @param {number} rowIndex - The 1-based row index of the target cell.
 * @param {number} columnIndex - The 1-based column index of the target cell.
 * @returns {Array<string>} - An array of email addresses extracted from the cell.
 */
function getEmailsFromSheetCell(sheetName, rowIndex, columnIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet '${sheetName}' not found!`);
    return [];
  }
  
  // Get the cell value directly. 
  // Note: rowIndex and columnIndex here are assumed to be 1-indexed.
  const cellValue = sheet.getRange(rowIndex, columnIndex).getValue();
  if (!cellValue) return [];
  
  // Split the comma-separated string into an array,
  // trim any extra whitespace, and filter out any empty strings.
  return cellValue
    .split(",")
    .map(email => email.trim())
    .filter(email => email);
}


/**
 * Exports a JSON object from this script to your Google Drive. 
 * From there, the file can be downloaded and the data can be analyzed with more powerful tools than what JavaScript's capabilities. 
 */
function exportStudentDataToDrive() {
  const jsonData = PropertiesService.getScriptProperties().getProperty("insert_JSON_object_name");
  const file = DriveApp.createFile("student_data.json", jsonData, MimeType.PLAIN_TEXT);
  Logger.log("File created: " + file.getUrl());
}



