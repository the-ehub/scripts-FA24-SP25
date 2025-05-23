/**
 * Builds and stores a structured student data object from interest and availability sheets.
 *
 * This function merges data from the "interest-responses" and "final-availabilities" sheets
 * into a single studentData object keyed by email. You can control which form is required
 * for a student to be included.
 *
 * @param {boolean} [requireAvailability=true] - If true, only includes students who submitted availability.
 * @param {boolean} [requireInterests=true] - If true, only includes students who submitted interest data.
 * @param {string} [propertyKey="studentData"] - Optional property key for storing the final object.
 * @returns {void}
 */
function preprocessStudentData(requireAvailability = false, requireInterests = true, propertyKey = "studentData") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const interestsSheet = ss.getSheetByName("interest-responses");
  const availSheet = ss.getSheetByName("final-availabilities");

  const studentData = {};

  // ---- Parse interest data ----
  let interestLookup = {};
  if (interestsSheet) {
    const interestData = interestsSheet.getDataRange().getValues();
    const headers = interestData[0];
    const emailIndex = headers.indexOf("Email");
    const firstNameIndex = headers.indexOf("First Name");
    const lastNameIndex = headers.indexOf("Last Name");
    const interestsIndex = headers.indexOf("Interests");

    if ([emailIndex, firstNameIndex, lastNameIndex, interestsIndex].includes(-1)) {
      Logger.log("Missing one or more required columns in interest-responses.");
      if (requireInterests) return;
    } else {
      for (let i = 1; i < interestData.length; i++) {
        const email = interestData[i][emailIndex]?.trim().toLowerCase();
        if (!email) continue;
        interestLookup[email] = {
          firstName: interestData[i][firstNameIndex],
          lastName: interestData[i][lastNameIndex],
          interests: interestData[i][interestsIndex]?.split(";").map(i => i.trim()) || []
        };
      }
    }
  }

  // ---- Parse availability data ----
  let availabilityLookup = {};
  if (availSheet) {
    const availData = availSheet.getDataRange().getValues();
    const headers = availData[0];
    const emailIndex = headers.indexOf("Email Address");
    const firstNameIndex = headers.indexOf("First Name");
    const lastNameIndex = headers.indexOf("Last Name");

    if ([emailIndex, firstNameIndex, lastNameIndex].includes(-1)) {
      Logger.log("Missing one or more required columns in final-availabilities.");
      if (requireAvailability) return;
    } else {
      const days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
      const availIndices = {};
      days.forEach(day => {
        availIndices[day] = headers.indexOf(`What times are you typically available for the Spring '25 semester? [${day}]`);
      });

      for (let i = 1; i < availData.length; i++) {
        const email = availData[i][emailIndex]?.trim().toLowerCase();
        if (!email) continue;

        let availability = {};
        days.forEach(day => {
          const idx = availIndices[day];
          if (idx !== -1) {
            const timeSlots = availData[i][idx]?.split(",").map(t => t.trim()).filter(t => t);
            if (timeSlots.length > 0) availability[day] = timeSlots;
          }
        });

        availabilityLookup[email] = {
          firstName: availData[i][firstNameIndex],
          lastName: availData[i][lastNameIndex],
          availability: availability
        };
      }
    }
  }

  // ---- Merge student data ----
  const allEmails = new Set([...Object.keys(interestLookup), ...Object.keys(availabilityLookup)]);
  allEmails.forEach(email => {
    const hasInterest = !!interestLookup[email];
    const hasAvailability = !!availabilityLookup[email];

    if ((requireInterests && !hasInterest) || (requireAvailability && !hasAvailability)) return;

    studentData[email] = {
      email,
      firstName: interestLookup[email]?.firstName || availabilityLookup[email]?.firstName || "",
      lastName: interestLookup[email]?.lastName || availabilityLookup[email]?.lastName || "",
      interests: interestLookup[email]?.interests || [],
      availability: availabilityLookup[email]?.availability || {}
    };
  });

  // ---- Store in Script Properties ----
  PropertiesService.getScriptProperties().setProperty(propertyKey, JSON.stringify(studentData));
  Logger.log(`Stored ${Object.keys(studentData).length} students in property: ${propertyKey}`);
}

/**
 * Logs a query for day, time, and interests and generates a unique output sheet name.
 *
 * @param {string} day - Day of the week (e.g., "Monday").
 * @param {string} time - Time slot (e.g., "10am-11am").
 * @param {Array<string>} interests - List of selected interests.
 * @returns {string} - A unique sheet name associated with the query.
 */
function logQuery(day, time, interests) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Query Log");
  
  // Create log sheet if it doesn't exist
  if (!logSheet) {
    logSheet = ss.insertSheet("Query Log");
    logSheet.appendRow(["Query ID", "Date", "Time", "Day", "Interests", "Sheet Name"]);
  }

  // Generate a unique query ID
  const queryId = "Q" + new Date().getTime().toString().slice(-5); // Shortened timestamp
  const sheetName = `${day}_${time}_${queryId}`;

  // Log the query
  logSheet.appendRow([queryId, new Date().toLocaleString(), time, day, interests.join("; "), sheetName]);

  return sheetName; // Return the name to create the actual sheet
}


/**
 * Processes a sidebar filter form submission and writes matched students
 * to a new sheet based on day, time, and interests.
 *
 * @param {Object} data - Object with keys: day, time, and interests.
 */ 
function processFilters(data) {
  const jsonData = PropertiesService.getScriptProperties().getProperty("studentData");
  if (!jsonData) {
    Logger.log("No preprocessed data found. Run preprocessStudentData() first.");
    return;
  }
  
  const studentData = JSON.parse(jsonData);
  const { day, time, interests } = data;

  let matchingStudents = {};

  if (!data.interests || data.interests.length === 0) { 
    matchingStudents = Object.values(studentData).filter(student => {
      return (
        student.availability[day] && student.availability[day].includes(time)
      );
    });
  } else {
    matchingStudents = Object.values(studentData).filter(student => {
      return (
        student.availability[day] && student.availability[day].includes(time) &&
        interests.some(interest => student.interests.includes(interest.toLowerCase()))
      );
    });
  }


  if (matchingStudents.length === 0) {
    Logger.log("No students found for the given filters.");
    return;
  }

  // Generate a sheet name using the query log
  const sheetName = logQuery(data.day, data.time, data.interests);

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  sheet.appendRow(["Email", "First Name", "Last Name", "Interests"]);
  matchingStudents.forEach(student => {
    sheet.appendRow([student.email, student.firstName, student.lastName, student.interests.join(", ")]);
  });

  Logger.log(`Filtered data written to sheet: ${sheetName}`);
}

