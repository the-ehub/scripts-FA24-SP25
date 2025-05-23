/**
 * Adds a membership column and status for each student in the "data-driven-matchmaking" sheet. 
 * "data-driven-matchmaking" sheet is pulled from a Hubspot list that contains all the people who opted in to have their information shared with matches and selected that they are either actively looking for people to work on their idea or for people with existing ideas to work with. 
 * Membership status is pulled from another sheet containing all the eHub members. 
 * All students whose membership status cannot be found in the member sheet are deleted from the data-driven-matchmaking sheet. 
 */
function updateMembershipColumn() {
  var matchmakingFile = SpreadsheetApp.getActiveSpreadsheet(); // Active file
  var matchmakingSheet = matchmakingFile.getSheetByName("data-driven-matchmaking");

  var ehubFile = SpreadsheetApp.openById("1iLMmkR2h0BhJys9Psnx6GXjPUBlNxGrhVOlGw7ONPO4"); // Open eHub S2025 Membership Sheet
  var ehubSheet = ehubFile.getSheetByName("all eHub members");

  // Get data from "data-driven-matchmaking"
  var matchmakingData = matchmakingSheet.getDataRange().getValues();
  var emailIndexMatchmaking = matchmakingData[0].indexOf("Email"); // Get email column index
  var membershipCol = matchmakingData[0].length; // New column for membership
  
  // Get data from "all eHub members"
  var ehubData = ehubSheet.getDataRange().getValues();
  var emailIndexEhub = ehubData[0].indexOf("Email");
  var membershipIndexEhub = ehubData[0].indexOf("Track");

  if (emailIndexMatchmaking == -1 || emailIndexEhub == -1 || membershipIndexEhub == -1) {
    Logger.log("Required columns not found.");
    return;
  }

  // Create a lookup dictionary for memberships
  var membershipMap = {};
  for (var i = 1; i < ehubData.length; i++) {
    var email = ehubData[i][emailIndexEhub].trim().toLowerCase();
    var membership = ehubData[i][membershipIndexEhub];
    if (email) {
      membershipMap[email] = membership;
    }
  }

  // Add "Membership" column header if not already there
  if (matchmakingData[0].indexOf("Membership") == -1) {
    matchmakingSheet.getRange(1, membershipCol + 1).setValue("Membership");
  }

  // Iterate from bottom up and remove rows with no membership
  for (var j = matchmakingData.length - 1; j > 0; j--) {
    var studentEmail = matchmakingData[j][emailIndexMatchmaking].trim().toLowerCase();
    var membership = membershipMap[studentEmail] || "";

    if (membership) {
      matchmakingSheet.getRange(j + 1, membershipCol + 1).setValue(membership);
    } else {
      matchmakingSheet.deleteRow(j + 1);
    }
  }
}

/**
 * Validates that the given sheet contains all required columns for teammate matching, order-agnostic.
 * 
 * @param {Sheet} sheet - The Google Sheet to validate.
 * @returns {Object|null} - A mapping of required column names to their indices if all are present; 
 *                          otherwise, returns null.
 */
function validateSheet(sheet) {
  const requiredColumns = [
    'Email', 'Skills Needed in Teammates', 'Skills to Contribute',
    'Interests', 'teammate_desribe_yourself', 'teammate_looking_for',
    'Seeking Teammates to Work on Your Idea?', 'Seeking Team to Join and Work on Their Idea?',
    'Membership'
  ];
  const headers = sheet.getDataRange().getValues()[0];
  const columnIndices = {};
  requiredColumns.forEach(col => {
    columnIndices[col] = headers.indexOf(col);
  });
  return Object.values(columnIndices).every(index => index !== -1) ? columnIndices : null;
}

/**
 * Separates students into leaders and members based on their membership type and interest in team roles.
 *
 * @param {Array<Array>} data - The full dataset (excluding headers) as an array of rows.
 * @param {Object} columnIndices - An object mapping column names to their indices (output from validateSheet()). 
 * @returns {Object} - An object with two arrays: `leaders` (students seeking teammates for their own idea)
 *                     and `members` (students seeking to join someone else's team).
 */
function separateLeadersAndMembers(data, columnIndices) {
  let leaders = [], members = [];
  data.forEach(row => {
    const email = row[columnIndices['Email']];
    const seekingTeammates = row[columnIndices['Seeking Teammates to Work on Your Idea?']];
    const seekingTeam = row[columnIndices['Seeking Team to Join and Work on Their Idea?']];
    const membership = row[columnIndices['Membership']];
    if ((membership === 'BUILD' || membership === 'BUILDdiscover') && seekingTeammates === 'Yes') {
      leaders.push(row);
    } else if (seekingTeam === 'Yes') {
      members.push(row);
    }
  });
  return { leaders, members };
}


/**
 * Finds the top 5 best-matching team members for a given team leader based on a match score.
 *
 * @param {Array} leader - A row representing the team leader's data.
 * @param {Array<Array>} members - An array of rows, each representing a potential team member.
 * @param {Object} columnIndices - An object mapping column names to their indices (output from validateSheet()).
 * @returns {Array<Object>} - An array of up to 5 objects, each containing a `member` row and its `score`.
 */
function findTopMatches(leader, members, columnIndices) {
  let matches = members.map(member => {
    const matchScore = calculateMatchScore(leader, member, columnIndices);
    return { member, score: matchScore };
  }).sort((a, b) => b.score - a.score);
  return matches.slice(0, 5);
}

/**
 * Calculates a compatibility score between a team leader and a potential member based on shared interests
 * and complementary skill needs.
 *
 * @param {Array} leader - A row representing the team leader's data.
 * @param {Array} member - A row representing the team member's data.
 * @param {Object} columnIndices - An object mapping column names to their indices.
 * @returns {number} - A numeric match score. Higher scores indicate better alignment.
 */
function calculateMatchScore(leader, member, columnIndices) {
  let score = 0;
  const commonInterests = getCommonItems(leader[columnIndices['Interests']], member[columnIndices['Interests']]);
  const leaderNeeds = getCommonItems(leader[columnIndices['Skills Needed in Teammates']], member[columnIndices['Skills to Contribute']]);
  const memberNeeds = getCommonItems(member[columnIndices['Skills Needed in Teammates']], leader[columnIndices['Skills to Contribute']]);
  if (commonInterests.length) score += commonInterests.length * 3;
  if (leaderNeeds.length) score += leaderNeeds.length * 2;
  if (memberNeeds.length) score += memberNeeds.length;
  return score;
}

/**
 * Returns the common items between two semicolon-separated strings.
 *
 * @param {string} list1 - A semicolon-separated string (e.g., "AI; healthcare; software").
 * @param {string} list2 - Another semicolon-separated string.
 * @returns {Array<string>} - An array of items present in both lists.
 */
function getCommonItems(list1, list2) {
  const set1 = new Set(list1.split(';').map(i => i.trim()));
  const set2 = new Set(list2.split(';').map(i => i.trim()));
  return [...set1].filter(item => set2.has(item));
}

/**
 * Writes all leader-member match results to a Google Sheet, including scores and match details.
 *
 * @param {Sheet} sheet - The sheet where match results will be written.
 * @param {Array<Object>} matches - An array of match objects with leader, member, score, and overlap info.
 * @param {Object} columnIndices - An object mapping column names to their indices in the data rows.
 */
function writeMatchesToSheet(sheet, matches, columnIndices) {
  const headers = [
    'Team Leader Email', 'Team Member Email', 'Match Score', 'Common Interests',
    'Skills Leader Needs', 'Skills Member Needs', 'Match Description',
    'Leader Description', 'What Member is Looking For', 'What Leader is Looking For'
  ];
  sheet.clear();
  sheet.appendRow(headers);
  matches.forEach(({ leader, member, score, commonInterests, leaderNeeds, memberNeeds }) => {
    sheet.appendRow([
      leader[columnIndices['Email']], member[columnIndices['Email']], score,
      commonInterests.join('; '), leaderNeeds.join('; '), memberNeeds.join('; '),
      member[columnIndices['teammate_desribe_yourself']], leader[columnIndices['teammate_desribe_yourself']],
      member[columnIndices['teammate_looking_for']], leader[columnIndices['teammate_looking_for']]
    ]);
  });
}

/**
 * Runs the full teammate matchmaking process:
 * - Validates the sheet
 * - Separates leaders and members
 * - Calculates top matches for each leader
 * - Writes all matches with scores and overlap details to the "Matches" sheet
 */
function runMatchmaking() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data-driven-matchmaking');
  const columnIndices = validateSheet(sheet);
  if (!columnIndices) {
    Logger.log('Sheet validation failed. Please check column names.');
    return;
  }
  const data = sheet.getDataRange().getValues().slice(1);
  const { leaders, members } = separateLeadersAndMembers(data, columnIndices);
  let allMatches = [];
  leaders.forEach(leader => {
    const topMatches = findTopMatches(leader, members, columnIndices);
    topMatches.forEach(({ member, score }) => {
      allMatches.push({
        leader, member, score,
        commonInterests: getCommonItems(leader[columnIndices['Interests']], member[columnIndices['Interests']]),
        leaderNeeds: getCommonItems(leader[columnIndices['Skills Needed in Teammates']], member[columnIndices['Skills to Contribute']]),
        memberNeeds: getCommonItems(member[columnIndices['Skills Needed in Teammates']], leader[columnIndices['Skills to Contribute']])
      });
    });
  });
  const matchSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matches') ||
                     SpreadsheetApp.getActiveSpreadsheet().insertSheet('Matches');
  writeMatchesToSheet(matchSheet, allMatches, columnIndices);
}

