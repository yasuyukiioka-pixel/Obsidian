/**
 * @file dataValidation.js
 * @brief This file contains functions for data validation, such as checking for duplicates and invalid email formats.
 */

/**
 * @const {RegExp} EMAIL_REGEX
 * @brief A regular expression for validating standard email address formats.
 */
const EMAIL_REGEX = /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;

/**
 * @function areEmailsValid
 * @brief Validates a string containing one or more email addresses.
 *
 * The string can contain multiple emails separated by newlines. Each email is trimmed
 * and checked against the EMAIL_REGEX. The function returns true only if all email
 * addresses in the string are valid.
 *
 * @param {string} emailString The string of email addresses to validate.
 * @returns {boolean} True if all emails are valid, false otherwise. An empty string is considered valid.
 */
function areEmailsValid(emailString) {
  if (!emailString || emailString.trim() === '') {
    return true;
  }
  const emails = emailString.split('\n').filter(e => e.trim() !== '');
  return emails.every(email => EMAIL_REGEX.test(String(email).trim().toLowerCase()));
}

/**
 * @function findInvalidEmails
 * @brief Scans a map of data for entries with improperly formatted email addresses.
 *
 * This function iterates through a map of team data and uses `areEmailsValid` to check
 * the 'to' and 'cc' fields. If an invalid email is found, it creates a formatted row
 * for the results report, indicating whether the issue is with the TO, CC, or both.
 *
 * @param {Map<string, {to: string, cc: string}>} updatedDataMap A map of the data to be validated.
 * @returns {Array<Array<string>>} A 2D array of rows corresponding to entries with invalid emails.
 */
function findInvalidEmails(updatedDataMap) {
  const invalidEmailRows = [];

  for (const [teamName, entry] of updatedDataMap.entries()) {
    const isToValid = areEmailsValid(entry.to);
    const isCcValid = areEmailsValid(entry.cc);

    if (!isToValid || !isCcValid) {
      let remarks = '不適切：';
      const issues = [];
      if (!isToValid) {
        issues.push('TO');
      }
      if (!isCcValid) {
        issues.push('CC');
      }
      remarks += issues.join('、');

      invalidEmailRows.push([
        teamName,
        entry.to,
        entry.cc,
        'M',
        remarks
      ]);
    }
  }
  return invalidEmailRows;
}


/**
 * @function findDuplicates
 * @brief Checks for duplicate team names within a given sheet.
 *
 * It works in two passes:
 * 1. It counts the occurrences of each team name.
 * 2. It iterates through the data again and flags any team name that appeared more than once.
 * This ensures that each duplicated team is reported only once.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to check for duplicates.
 * @returns {Array<Array<string>>} A 2D array of rows for any teams that were found to be duplicated.
 * @throws {Error} If the required 'Team Name' column is not found in the sheet.
 */
function findDuplicates(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const teamNameIndex = headers.indexOf(COLUMN_DEF.TEAM_NAME);
  const toIndex = headers.indexOf(COLUMN_DEF.TO);
  const ccIndex = headers.indexOf(COLUMN_DEF.CC);

  if (teamNameIndex === -1) {
    throw new Error(`重複チェックエラー: 「${COLUMN_DEF.TEAM_NAME}」列が見つかりません。`);
  }

  const teamCounts = new Map();
  const duplicateRows = [];

  for (let i = 1; i < data.length; i++) {
    const teamName = data[i][teamNameIndex];
    if (teamName) {
      const trimmedName = String(teamName).trim();
      teamCounts.set(trimmedName, (teamCounts.get(trimmedName) || 0) + 1);
    }
  }

  const processedDuplicates = new Set();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rawTeamName = row[teamNameIndex];

    if (!rawTeamName) continue;

    const trimmedName = String(rawTeamName).trim();
    const count = teamCounts.get(trimmedName);

    if (count > 1 && !processedDuplicates.has(trimmedName)) {
      duplicateRows.push([
        trimmedName,
        row[toIndex] || '',
        row[ccIndex] || '',
        '重複',
        `${count}回出現`
      ]);
      processedDuplicates.add(trimmedName);
    }
  }

  return duplicateRows;
}

/**
 * @function identifyNewTeamDuplicates
 * @brief Identifies teams in the source data that already exist in the master data.
 *
 * @param {Array<Array<any>>} sourceData Raw data rows from source sheet.
 * @param {Map<string, object>} masterDataMap Map of existing teams from master sheet.
 * @param {number} teamNameIndex Column index (0-based) where the Team Name is located in sourceData.
 * @returns {Array<object>} List of duplicate entries { teamName, sourceRowIndex }.
 */
function identifyNewTeamDuplicates(sourceData, masterDataMap, teamNameIndex) {
  const duplicates = [];

  if (teamNameIndex < 0) {
    Logger.log('identifyNewTeamDuplicates: Invalid teamNameIndex provided.');
    return [];
  }

  const normalize = (str) => String(str)
    .normalize('NFKC')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .trim();

  sourceData.forEach((row, index) => {
    // Check if the row has enough columns
    if (row.length <= teamNameIndex) return;

    const rawTeamName = row[teamNameIndex];
    // Ignore empty cells
    if (!rawTeamName || String(rawTeamName).trim() === '') return;

    const teamName = normalize(rawTeamName);

    // Skip possible header values if they match the column definition
    if (teamName === 'チーム名' || teamName === COLUMN_DEF.TEAM_NAME) return;

    if (masterDataMap.has(teamName)) {
      duplicates.push({
        teamName: teamName,
        sourceRowIndex: index + 1 // +1 for 1-based row reference (assuming sourceData starts from row 1 relative to range)
                                  // Caller must adjust if sourceData is just a subset.
                                  // Let's assume index is relative to the array passed.
      });
    }
  });

  return duplicates;
}

/**
 * Checks for duplicates between new team names and the master team list.
 * Supports both Exact Match and Partial Match.
 * * @param {Array<object>} newRecords - Array of objects { teamName, rowNum, ... }
 * @param {Array<string>} masterTeams - Array of existing team names from the master list.
 * @returns {Array<object>} Array of result objects { type: '完全'|'部分', teamName, rowNums: [], matchTarget: string }
 */
function checkTeamNameDuplicates(newRecords, masterTeams) {
  const results = [];
  const processedTeams = new Set(); // To group by team name if same team appears multiple times in newRecords

  // Helper to normalize strings for comparison (remove spaces, etc? User didn't specify, but safer)
  // For now, let's stick to simple trim() as requested "Partial Match" might rely on spaces.

  newRecords.forEach(record => {
    const newTeam = record.teamName;
    const rowNum = record.rowNum;

    // Check against Master List
    let exactMatchFound = false;
    let partialMatches = [];

    masterTeams.forEach(masterTeam => {
      if (newTeam === masterTeam) {
        exactMatchFound = true;
      } else if (newTeam.includes(masterTeam) || masterTeam.includes(newTeam)) {
        partialMatches.push(masterTeam);
      }
    });

    if (exactMatchFound) {
      addResult(results, '完全', newTeam, rowNum, newTeam); // Target is itself
    } else if (partialMatches.length > 0) {
      // For partial matches, we might have multiple candidates.
      // Let's record all of them or just the first?
      // User format: "X社、100：200" implies grouping by the NEW team.
      // matchTarget might be useful to show WHAT it matched against.
      addResult(results, '部分', newTeam, rowNum, partialMatches.join(', '));
    }
  });

  return results;
}

/**
 * Helper to add or update results in the list.
 * Groups by (Type + TeamName).
 */
function addResult(results, type, teamName, rowNum, matchTarget) {
  const existing = results.find(r => r.type === type && r.teamName === teamName);
  if (existing) {
    if (!existing.rowNums.includes(rowNum)) {
      existing.rowNums.push(rowNum);
    }
    // Update matchTarget if new ones found? (For partial)
    if (type === '部分' && !existing.matchTarget.includes(matchTarget)) {
       existing.matchTarget += `, ${matchTarget}`;
    }
  } else {
    results.push({
      type: type,
      teamName: teamName,
      rowNums: [rowNum],
      matchTarget: matchTarget
    });
  }
}