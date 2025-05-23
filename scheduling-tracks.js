/**
 * Retrieves all emails from a specified column in a given sheet.
 *
 * @param {string} sheetName - The name of the sheet to read from.
 * @param {number} columnIndex - The index of the column containing email addresses.
 * @returns {Array<string>} - A list of emails found in the specified column.
 */
function getEmailsFromSheet(sheetName, columnIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet '${sheetName}' not found!`);
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => row[columnIndex]).filter(email => email);
}

/**
 * Retrieves emails from a specified column where the corresponding value
 * in another column is "Yes".
 *
 * @param {string} sheetName - The name of the sheet to read from.
 * @param {number} emailColumnIndex - The index of the column containing email addresses.
 * @param {number} conditionColumnIndex - The index of the column containing "Yes"/"No" values.
 * @returns {Array<string>} - A list of emails where the condition is "Yes".
 */
function getEmailsWithCondition(sheetName, emailColumnIndex, conditionColumnIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('Sheet not found: ' + sheetName);
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  return data.slice(1) // Skip headers
    .filter(row => row[conditionColumnIndex] === "Yes") // Keep only rows where the condition column is "Yes"
    .map(row => row[emailColumnIndex]) // Extract emails
    .filter(email => email); // Remove empty values
}

/**
 * Returns emails of students in a specified track from the "final-availabilities" sheet.
 * Handles students listed under multiple tracks (comma-separated).
 *
 * @param {string} track - The name of the track to filter by.
 * @returns {Array<string>} - A list of emails of students in the given track.
 */ 
function getEmailsByTrack(track) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("final-availabilities");
  if (!sheet) {
    Logger.log("Sheet 'final-availabilities' not found!");
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const trackIndex = 4;
  const emailIndex = 1;
  
  return data.slice(1).reduce((emails, row) => {
    if (row[trackIndex].split(',').map(t => t.trim()).includes(track)) {
      emails.push(row[emailIndex]);
    }
    return emails;
  }, []);
}

/**
 * Finds time slots for a given track where at least a specified number of students are available.
 *
 * @param {string} track - The name of the track to filter students by.
 * @param {number} thresh - Minimum number of students required for a time slot to be included.
 * @returns {[Object, Array<string>]} - A two-element array:
 *   - First: a dictionary with time slots as keys and lists of student emails as values.
 *   - Second: a list of students who did not fit into any qualifying time slot.
 */
function findAvailableSlots(track, thresh) { 

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("final-availabilities");
  const data = sheet.getDataRange().getValues();

  const trackIndex = 4;
  const emailIndex = 1;
  const availabilityIndices = [5, 6, 7, 8, 9, 10, 11]; // Assuming availability columns are from C (3) to I (9)
  const days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

  const availableSlots = {};
  const studentsInTrack = [];
  const excludedStudents = [];

  // Collect students in the specified track and their availabilities
  for (let i = 1; i < data.length; i++) {
    const studentTracks = data[i][trackIndex].split(',').map(t => t.trim()); // Split by comma
    const found = studentTracks.some(string => {
      if (string.includes(track)) {
        return true; // Stops the loop on the first match
      }
      return false; // Continue the loop
    });
    const email = data[i][emailIndex];

    if (found) {
      studentsInTrack.push(email);
      // Check each availability column
      availabilityIndices.forEach((index) => {
        const timeSlotsString = data[i][index];
        if (timeSlotsString) { // Check if the availability column is not empty
          const timeSlots = timeSlotsString.split(',').map(slot => slot.trim());
          timeSlots.forEach(slot => {
            const key = `${days[index-5]} ${slot}`;
            if (!availableSlots[key]) {
              availableSlots[key] = [];
            }
            availableSlots[key].push(email);
          });
        }
      });
    }
  }

  // Filter out slots with less than THRESH students
  const filteredSlots = {};
  for (const [key, emails] of Object.entries(availableSlots)) {
    if (emails.length >= thresh) {
      filteredSlots[key] = emails;
    }
  }

  // Check if all students are included
  const includedStudents = new Set();
  for (const emails of Object.values(filteredSlots)) {
    emails.forEach(email => includedStudents.add(email));
  }

  studentsInTrack.forEach(email => {
    if (!includedStudents.has(email)) {
      excludedStudents.push(email);
    }
  });

  return [filteredSlots, excludedStudents];
  
}

/**
 * Finds time slots where at least a threshold number of provided students are available.
 *
 * @param {Array<string>} emails - List of student emails to consider.
 * @param {number} thresh - Minimum number of students required per time slot.
 * @returns {[Object, Array<string>]} - A two-element array:
 *   - First: a dictionary with time slots as keys and lists of student emails as values.
 *   - Second: a list of students who did not fit into any qualifying time slot.
 */
function findAvailableSlotsForEmails(emails, thresh) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("final-availabilities");
  const data = sheet.getDataRange().getValues();
  
  const emailIndex = 1;
  const availabilityIndices = [5, 6, 7, 8, 9, 10, 11];
  const days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
  
  const availableSlots = {};
  const excludedStudents = [];

  for (let i = 1; i < data.length; i++) {
    const email = data[i][emailIndex];
    if (!emails.includes(email)) continue;
    
    availabilityIndices.forEach((index) => {
      const timeSlotsString = data[i][index];
      if (timeSlotsString) {
        const timeSlots = timeSlotsString.split(',').map(slot => slot.trim());
        timeSlots.forEach(slot => {
          const key = `${days[index - 5]} ${slot}`;
          if (!availableSlots[key]) availableSlots[key] = [];
          availableSlots[key].push(email);
        });
      }
    });
  }

  const filteredSlots = Object.fromEntries(
    Object.entries(availableSlots).filter(([_, list]) => list.length >= thresh)
  );

  const includedStudents = new Set(Object.values(filteredSlots).flat());
  excludedStudents.push(...emails.filter(email => !includedStudents.has(email)));

  return [filteredSlots, excludedStudents];
}

/**
 * Selects the minimum number of time slots needed to cover all students at least once.
 *
 * @param {Object} filteredSlots - A dictionary of time slots to student email lists.
 * @returns {Object} - A reduced dictionary of time slots covering all students at least once.
 */
function getMinimumTimeSlots(filteredSlots) {
  const studentsCovered = new Set(); // Track students who are already covered
  const selectedSlots = {}; // Store selected time slots with corresponding emails

  // Get all unique students
  const allStudents = new Set(Object.values(filteredSlots).flat());

  // While there are still students not covered
  while (studentsCovered.size < allStudents.size) {
    
    // Convert the object into a sorted array of entries based on new coverage
    const sortedSlots = Object.entries(filteredSlots).sort((a, b) => {
      const newCoverageA = a[1].filter(email => !studentsCovered.has(email)).length;
      const newCoverageB = b[1].filter(email => !studentsCovered.has(email)).length;
      return newCoverageB - newCoverageA; // Sort descending by new coverage
    });

    // Pick the first slot from the sorted list (it has the most new coverage)
    const [bestSlotKey, bestSlotEmails] = sortedSlots[0];

    if (bestSlotEmails) {
      // Add the best slot and its corresponding emails to selectedSlots
      selectedSlots[bestSlotKey] = bestSlotEmails;
      // Update the covered students
      bestSlotEmails.forEach(email => studentsCovered.add(email));
      // Remove the selected slot from the filteredSlots to avoid reconsidering it
      delete filteredSlots[bestSlotKey];
    } else {
      break; // Break if no new coverage is found
    }
  }

  return selectedSlots; // Return the dictionary of selected slots with emails
}

/**
 * Selects popular time slots iteratively until all students are covered,
 * and writes results directly to a sheet.
 *
 * @param {Object} selectedSlots - A dictionary of time slots to student emails.
 * @param {number} minGroupSize - Minimum number of students required in a slot.
 * @param {string} sheetName - Name of the output sheet to write to.
 * @param {Array<string>} [virtualStudents=[]] - Optional list of virtual student emails.
 * @param {Array<string>} [excludedStudents=[]] - Optional list of excluded student emails.
 */ 
function getIterativeTimeSlots(selectedSlots, minGroupSize, sheetName, virtualStudents = [], excludedStudents = []) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  if (!sheet) {
    Logger.log('Sheet with name ' + sheetName + ' not found!');
    return;
  }
  
  // Clear the sheet to start fresh
  sheet.clear();
  
  // Write headers in the first row
  const headers = ['Time Slot', 'Students', 'Num Students', 'Excluded Students', 'Num Excluded Students', 'Total Students'];
  sheet.appendRow(headers);
  
  let remainingStudents = new Set();
  Object.values(selectedSlots).forEach(studentList => {
    studentList.forEach(student => remainingStudents.add(student));
  });

  let finalSelections = [];

  let totalStudents = 0;

  while (remainingStudents.size >= minGroupSize) {
    // Rank slots by the number of available students
    let sortedSlots = Object.entries(selectedSlots)
      .map(([timeSlot, studentEmails]) => {
        const filteredEmails = studentEmails.filter(email => remainingStudents.has(email));
        return [timeSlot, filteredEmails, filteredEmails.length];
      })
      .filter(([_, __, count]) => count > 0) // Only keep non-empty slots
      .sort((a, b) => b[2] - a[2]); // Sort by student count (descending)

    if (sortedSlots.length === 0) break; // Stop if no valid time slots are left

    let [chosenSlot, chosenStudents, count] = sortedSlots[0]; // Pick the slot with most students

    if (count < minGroupSize) break; // Stop if the group size is below the threshold

    finalSelections.push([chosenSlot, chosenStudents.join(', '), count]);

    // Remove selected students from the remaining pool
    chosenStudents.forEach(student => remainingStudents.delete(student));
    totalStudents += count;
  }

  // Write results to the sheet
  finalSelections.forEach(row => sheet.appendRow(row));

  if (excludedStudents.length > 0) {
    sheet.getRange("D2").setValue(excludedStudents.join(', '));
  }
  
  sheet.getRange("E2").setValue(excludedStudents.length);
  sheet.getRange("F2").setValue(excludedStudents.length + totalStudents)

  
  Logger.log('Final selected slots written to sheet ' + sheetName);
}

/**
 * Writes selected time slots and availability details to a Google Sheet.
 *
 * @param {Object} params - Object of named parameters.
 * @param {Object} params.selectedSlots - Dictionary of time slots and associated student emails.
 * @param {string} params.sheetName - Name of the sheet to write to (created or cleared if exists).
 * @param {Array<string>} [params.virtualStudents=[]] - Optional list of emails for virtual students.
 * @param {Array<string>} [params.excludedStudents=[]] - Optional list of students excluded from all time slots.
 */
function writeSelectedSlotsToSheet({ selectedSlots, sheetName, virtualStudents = [], excludedStudents = [] }) {
  // Open/insert the Google Sheet by name
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  if (!sheet) {
    Logger.log('Sheet with name ' + sheetName + ' not found!');
    return;
  }
  
  // Clear the sheet to start fresh
  sheet.clear();
  
  // Write headers in the first row
  const headers = ['Time Slot', 'Student Emails', 'Num Students', 'Virtual Students','Exc Students'];
  sheet.appendRow(headers);
  

  // Convert selectedSlots object to an array and sort by numStudents (descending)
  const sortedSlots = Object.entries(selectedSlots)
    .map(([timeSlot, studentEmails]) => {
      // Find virtual students who are in this row's student emails
      const matchingVirtualStudents = virtualStudents.filter(student => studentEmails.includes(student)).join(', ');

      return [timeSlot, studentEmails.join(', '), studentEmails.length, matchingVirtualStudents];
    })
    .sort((a, b) => b[2] - a[2]); // Sort by "Num Students" in descending order

  // Write sorted rows to the sheet
  sortedSlots.forEach(row => sheet.appendRow(row));

  if (excludedStudents.length > 0) {
    sheet.getRange("E2").setValue(`${excludedStudents.join(', ')}`); // Write log message to correct cell
  } 
  
  Logger.log('Selected slots written to sheet ' + sheetName);
}

/**
 * Finds and writes student availabilities for a specific track to a new sheet.
 *
 * @param {string} track - Track name to filter students by.
 * @param {number} thresh - Minimum number of students per time slot.
 * @param {string} outputSheetName - Name of the sheet to write results to.
 * @param {number} [virtualIndex=-1] - Optional index from the track responses sheet
 *   indicating which students are virtual.
 */
function processAvailabilityByTrack(track, thresh, outputSheetName, virtualIndex=-1) {
  const emails = getEmailsByTrack(track);
  let virtualStudents = []
  if (virtualIndex >= 0) virtualStudents = getEmailsWithCondition(`${track}-responses`, 0, virtualIndex);
  const [allSlots, excStudents] = findAvailableSlotsForEmails(emails, thresh);
  const filtered = getMinimumTimeSlots(allSlots);
  writeSelectedSlotsToSheet({
    selectedSlots: filtered,
    sheetName: outputSheetName,
    virtualStudents: virtualStudents,
    excludedStudents: excStudents
  });
}

/**
 * Finds and writes student availabilities from a custom list of emails in a sheet.
 *
 * @param {string} sheetName - Name of the sheet containing student emails.
 * @param {number} columnIndex - Index of the column with email addresses.
 * @param {number} thresh - Minimum number of students per time slot.
 * @param {string} outputSheetName - Name of the sheet to write results to.
 * @param {number} [virtualIndex=-1] - Optional index indicating virtual students in track responses sheet.
 */
function processAvailabilityBySheet(sheetName, columnIndex, thresh, outputSheetName, virtualIndex=-1) {
  const emails = getEmailsFromSheet(sheetName, columnIndex);
  let virtualStudents = []
  if (virtualIndex >= 0) virtualStudents = getEmailsWithCondition(sheetName, columnIndex, virtualIndex);
  const [allSlots, excStudents] = findAvailableSlotsForEmails(emails, thresh);
  const filtered = getMinimumTimeSlots(allSlots);
  writeSelectedSlotsToSheet({
    selectedSlots: filtered,
    sheetName: outputSheetName,
    virtualStudents: virtualStudents,
    excludedStudents: excStudents
  });
}

/**
 * Examples of processAVailabilityBySheet() and getIterativeTimeSlots() usages.
 */
function main() {
  // processAvailabilityBySheet("BUILD-regular-responses", 0, 0, "BUILD-regular-schedule", virtualIndex = 7)
  // processAvailabilityBySheet("BUILD-discover-responses", 0, 0, "BUILD-discover-schedule", virtualIndex = 7)
  // processAvailabilityByTrack("SEARCH", 20, "SEARCH-schedule")
  // processAvailabilityByTrack("TEST", 20, "TEST-schedule", virtualIndex=7)
  // processAvailabilityBySheet("SEARCH-responses", 0, 0, "SEARCH-schedule")
  // processAvailabilityBySheet("TEST-responses", 0, 0, "TEST-schedule", virtualIndex = 7)
  // processAvailabilityBySheet("last-BUILD-regular-responses", 0, 0, "last-BUILD-regular-schedule", virtualIndex = 7)
  // processAvailabilityBySheet("dinner-party", 1, 0, "dinner-times", virtualIndex = -1)

  let emails = getEmailsFromSheet("SEARCH-responses", 0);
  let [allSlots, excStudents] = findAvailableSlotsForEmails(emails, 0);
  getIterativeTimeSlots(allSlots, 0, "SEARCH-schedule", virtualStudents = [], excludedStudents = excStudents)
  emails = getEmailsFromSheet("TEST-responses", 0);
  [allSlots, excStudents] = findAvailableSlotsForEmails(emails, 0);
  getIterativeTimeSlots(allSlots, 0, "TEST-schedule", virtualStudents = [], excludedStudents = excStudents)
  emails = getEmailsFromSheet("BUILD-regular-responses", 0);
  [allSlots, excStudents] = findAvailableSlotsForEmails(emails, 0);
  getIterativeTimeSlots(allSlots, 0, "BUILD-regular-schedule", virtualStudents = [], excludedStudents = excStudents)
  emails = getEmailsFromSheet("BUILD-discover-responses", 0);
  [allSlots, excStudents] = findAvailableSlotsForEmails(emails, 0);
  getIterativeTimeSlots(allSlots, 0, "BUILD-discover-schedule", virtualStudents = [], excludedStudents = excStudents)
  emails = getEmailsFromSheet("last-BUILD-regular-responses", 0);
  [allSlots, excStudents] = findAvailableSlotsForEmails(emails, 0);
  getIterativeTimeSlots(allSlots, 0, "last-BUILD-regular-schedule", virtualStudents = [], excludedStudents = excStudents)
}


