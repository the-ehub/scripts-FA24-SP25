/**
 * Generates individual Google Docs for each team leader summarizing their top matches.
 * Each document includes details about matched team members, including interests and skills.
 * Document links are recorded in a sheet titled "Match Docs Links".
 *
 * @param {string} infoSheet - The name of the sheet containing full student info (e.g., names, majors, etc.).
 * @param {string} matchSheet - The name of the sheet containing match results (leader-member pairs).
 */
function createMatchDocs(infoSheet, matchSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(matchSheet);
  const infosheet = ss.getSheetByName(infoSheet);
  const data = sheet.getDataRange().getValues();
  const studentData = infosheet.getDataRange().getValues();
  const batchSize = 5; //number of matches
  
  const folder = createFolderIfNotExists('Matches Folder Sp25');  // Folder path

  // Create or get the "Match Docs Links" sheet
  let linksSheet = ss.getSheetByName("Match Docs Links");
  if (!linksSheet) {
    linksSheet = ss.insertSheet("Match Docs Links");
  } else {
    linksSheet.clear(); // Clears all data from the sheet
  }

  linksSheet.appendRow(["Leader Email", "First Name", "Document Link"]);


  let linkEntries = [];

  
  for (let i = 1; i < data.length; i += batchSize) {
    let batch = data.slice(i, i + batchSize);
    
    let leaderEmail = batch[0][0]; // Leader email is the same for all 5 rows
    let matchInfo = {
      leaderEmail: leaderEmail,
      matches: []
    };
    
    batch.forEach(row => {
      matchInfo.matches.push({
        matchEmail: row[1],
        score: row[2],
        commonInterests: row[3],
        skillsLeaderNeeds: row[4],
        skillsMemberNeeds: row[5],
        matchDescription: row[6],
        matchLookingFor: row[8]
      });
    });
      
    const leaderName = getNameByEmail(leaderEmail, studentData);
    
    // Create a Google Doc for each team leader
    const doc = DocumentApp.create(`${leaderName[0]} ${leaderName[1]} Match Info`);
    const docBody = doc.getBody();
    const docId = doc.getId(); 
    const docUrl = `https://docs.google.com/document/d/${docId}`;


    // Add header for the team leader
    docBody.appendParagraph(`Match Information for Team Leader: ${leaderName[0]} ${leaderName[1]}`)
          .setHeading(DocumentApp.ParagraphHeading.HEADING1);

    // Add each match's information to the document
    matchInfo.matches.forEach(match => {
      const matchRow = getRowByEmail(match.matchEmail, studentData);  // Retrieve match details by email
      if (matchRow) {
        const firstName = matchRow[1];  
        const lastName = matchRow[2];  
        const gradYear = matchRow[7];  
        const major = matchRow[8];  
        const minor = matchRow[9];
        const studentClass = matchRow[15];

        // Add match details
        const bold = docBody.appendParagraph(`Match: ${firstName} ${lastName} (${match.matchEmail})`);
        bold.editAsText().setBold(true);
        const normal = docBody.appendParagraph(`Match Score: ${match.score}`);
        normal.editAsText().setBold(false);

        docBody.appendParagraph(`Graduation Year: ${gradYear}`);
        docBody.appendParagraph(`Student Classification: ${studentClass}`);
        docBody.appendParagraph(`Major: ${major}`);
        if (minor != '') {
          docBody.appendParagraph(`Minor: ${minor}`);
        }
        

        let paragraph = docBody.appendParagraph('');
        paragraph.appendText('Common interests: ').setBold(true);  
        match.commonInterests.split(';').map(word => word.trim()).forEach(interest => {
          docBody.appendListItem(interest).setGlyphType(DocumentApp.GlyphType.BULLET);
        });
        
        paragraph = docBody.appendParagraph('');
        paragraph.appendText(`Skills ${firstName} has that you need: `).setBold(true);
        match.skillsLeaderNeeds.split(';').map(word => word.trim()).forEach(skillNeed => {
          docBody.appendListItem(skillNeed).setGlyphType(DocumentApp.GlyphType.BULLET);
        });
        
        paragraph = docBody.appendParagraph('');
        paragraph.appendText(`Skills you have that ${firstName} wants: `).setBold(true);
        match.skillsMemberNeeds.split(';').map(word => word.trim()).forEach(skillWant => {
          docBody.appendListItem(skillWant).setGlyphType(DocumentApp.GlyphType.BULLET);
        });

        docBody.appendParagraph(`Who ${firstName} is:`).setBold(true);  
        docBody.appendListItem(match.matchDescription.replace(/^"(.*)"$/, '$1')).setGlyphType(DocumentApp.GlyphType.BULLET).setBold(false);

        docBody.appendParagraph(`What ${firstName} is looking for:`).setBold(true);  
        docBody.appendListItem(match.matchLookingFor.replace(/^"(.*)"$/, '$1')).setGlyphType(DocumentApp.GlyphType.BULLET).setBold(false);
        
        docBody.appendParagraph('\n-----------------\n');
      }
    });
    
    doc.saveAndClose();  
    
    // Move the document to the specified folder
    const docFile = DriveApp.getFileById(docId);  
    docFile.moveTo(folder); 

    // Store leader email, name, and doc link in an array for batch writing
    linkEntries.push([leaderEmail, leaderName[0], docUrl]); 

    Logger.log(`Created doc for Team Leader: ${leaderName[0]} ${leaderName[1]}`)
  }

  // Append all collected rows to the "Match Docs Links" sheet at once
  if (linkEntries.length > 0) {
    linksSheet.getRange(linksSheet.getLastRow() + 1, 1, linkEntries.length, 3).setValues(linkEntries);
  }

}

/**
 * Generates Google Docs for each team member listing their top-ranked team leader matches.
 * Each document includes detailed information about the leaders and compatibility insights.
 * Document links are recorded in a sheet titled "Member Match Docs Links".
 *
 * @param {string} infoSheet - The name of the sheet containing student profile data.
 * @param {string} matchSheet - The name of the sheet containing match results.
 */
function createMemberMatchDocs(infoSheet, matchSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(matchSheet);
  const infosheet = ss.getSheetByName(infoSheet);
  const data = sheet.getDataRange().getValues();
  const studentData = infosheet.getDataRange().getValues();
  
  const folder = createFolderIfNotExists('Member Matches Folder Sp25');  // Folder path

  // Create or get the "Member Match Docs Links" sheet
  let linksSheet = ss.getSheetByName("Member Match Docs Links");
  if (!linksSheet) {
    linksSheet = ss.insertSheet("Match Docs Links");
  } else {
    linksSheet.clear(); // Clears all data from the sheet
  }

  linksSheet.appendRow(["Leader Email", "First Name", "Document Link"]);

  let linkEntries = [];

  let memberMatches = {};

  // Group matches by team member email
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let memberEmail = row[1]; // Team member email
    let matchDetails = {
      leaderEmail: row[0],
      score: row[2],
      commonInterests: row[3],
      skillsLeaderNeeds: row[4],
      skillsMemberNeeds: row[5],
      leaderDescription: row[7],
      leaderLookingFor: row[9]
    };

    if (!memberMatches[memberEmail]) {
      memberMatches[memberEmail] = [];
    }
    memberMatches[memberEmail].push(matchDetails);

    // Sort matches for this member by score in descending order
    memberMatches[memberEmail].sort((a, b) => b.score - a.score);
  }


  // Create documents for each team member
  for (let memberEmail in memberMatches) {
    const memberName = getNameByEmail(memberEmail, studentData);
    
    // Create a Google Doc for each team member
    const doc = DocumentApp.create(`${memberName[0]} ${memberName[1]} Team Matches`);
    const docBody = doc.getBody();
    const docId = doc.getId(); 
    const docUrl = `https://docs.google.com/document/d/${docId}`;

    // Add header for the team member
    docBody.appendParagraph(`Match Information for Team Member: ${memberName[0]} ${memberName[1]}`)
          .setHeading(DocumentApp.ParagraphHeading.HEADING1);

    let matches = memberMatches[memberEmail];

    // Add each team leader's information to the document
    matches.forEach(match => {
      const leaderRow = getRowByEmail(match.leaderEmail, studentData);
      if (leaderRow) {
        const firstName = leaderRow[1];
        const lastName = leaderRow[2];
        const gradYear = leaderRow[7];
        const major = leaderRow[8];
        const minor = leaderRow[9];
        const studentClass = leaderRow[15];

        // Add leader details
        const bold = docBody.appendParagraph(`Team Leader: ${firstName} ${lastName} (${match.leaderEmail})`);
        bold.editAsText().setBold(true);
        const normal = docBody.appendParagraph(`Match Score: ${match.score}`);
        normal.editAsText().setBold(false);

        docBody.appendParagraph(`Graduation Year: ${gradYear}`);
        docBody.appendParagraph(`Student Classification: ${studentClass}`);
        docBody.appendParagraph(`Major: ${major}`);
        if (minor != '') {
          docBody.appendParagraph(`Minor: ${minor}`);
        }

        let paragraph = docBody.appendParagraph('');
        paragraph.appendText('Common interests: ').setBold(true);
        match.commonInterests.split(';').map(word => word.trim()).forEach(interest => {
          docBody.appendListItem(interest).setGlyphType(DocumentApp.GlyphType.BULLET);
        });

        paragraph = docBody.appendParagraph('');
        paragraph.appendText(`Skills you have that ${firstName} needs: `).setBold(true);
        match.skillsLeaderNeeds.split(';').map(word => word.trim()).forEach(skillNeed => {
          docBody.appendListItem(skillNeed).setGlyphType(DocumentApp.GlyphType.BULLET);
        });

        paragraph = docBody.appendParagraph('');
        paragraph.appendText(`Skills ${firstName} has that you want: `).setBold(true);
        match.skillsMemberNeeds.split(';').map(word => word.trim()).forEach(skillWant => {
          docBody.appendListItem(skillWant).setGlyphType(DocumentApp.GlyphType.BULLET);
        });

        docBody.appendParagraph(`Who ${firstName} is:`).setBold(true);
        docBody.appendListItem(match.leaderDescription.replace(/^"(.*)"$/, '$1'))
               .setGlyphType(DocumentApp.GlyphType.BULLET).setBold(false);


        docBody.appendParagraph(`What ${firstName} is looking for:`).setBold(true);
        docBody.appendListItem(match.leaderLookingFor.replace(/^"(.*)"$/, '$1'))
               .setGlyphType(DocumentApp.GlyphType.BULLET).setBold(false);

        docBody.appendParagraph('\n-----------------\n');
      }
    });

    doc.saveAndClose();  

    // Move the document to the specified folder
    const docFile = DriveApp.getFileById(docId);
    docFile.moveTo(folder);

    // Store member email, name, and doc link in an array for batch writing
    linkEntries.push([memberEmail, memberName[0], docUrl]);

    Logger.log(`Created doc for Team Member: ${memberName[0]} ${memberName[1]}`);
  }

  // Batch write document links to the sheet
  if (linkEntries.length > 0) {
    linksSheet.getRange(linksSheet.getLastRow() + 1, 1, linkEntries.length, 3).setValues(linkEntries);
  }
}


/**
 * Creates and returns a Google Drive folder with the given name.
 * If a folder with that name already exists, it returns the existing one.
 *
 * @param {string} folderName - The name of the folder to create or find.
 * @returns {GoogleAppsScript.Drive.Folder} - The created or existing folder.
 */
function createFolderIfNotExists(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();  // Return existing folder
  }
  return DriveApp.createFolder(folderName);  // Create new folder
}

/**
 * Runs the document generation process for both team leaders and team members.
 * Creates folders, generates Google Docs summarizing matches, and logs the links.
 * Assumes 'Matches' and 'data-driven-matchmaking' sheets exist in the spreadsheet.
 */
function docsMain() {
  infoSheet = 'data-driven-matchmaking';
  matchSheet = 'Matches';
  createMatchDocs(infoSheet, matchSheet);
  createMemberMatchDocs(infoSheet, matchSheet)

}

