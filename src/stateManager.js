/**
 * @file stateManager.js
 * @brief Manages the state of the checking process using the History Sheet.
 */

/**
 * Reads the last processed V-Cube Order Number from the History Sheet.
 * Checks the second row (first data row) for the latest entry.
 * * @returns {string} The last processed order number, or empty string if not found.
 */
function getLastProcessedOrderNumber() {
  const sheet = getHistorySheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return ''; // Only header exists
  
  // Reading B2 (Order Number in 2nd column of 2nd row)
  const value = sheet.getRange("B2").getValue();
  return String(value).trim();
}

/**
 * Updates the last processed V-Cube Order Number by inserting a new row in the History Sheet.
 * * @param {string} orderNumber - The new order number to save.
 * @param {string} teamName - The team name associated with the order number.
 */
function updateLastProcessedOrderNumber(orderNumber, teamName) {
  if (!orderNumber) return;
  const sheet = getHistorySheet();
  
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  
  // Insert new row after header (Row 2)
  sheet.insertRowAfter(1);
  
  // Set values: Date, Order Number, Team Name
  sheet.getRange(2, 1, 1, 3).setValues([[today, orderNumber, teamName]]);
}

/**
 * Helper to get the History Sheet. Creates it if it doesn't exist.
 */
function getHistorySheet() {
  const ss = SpreadsheetApp.openById(CONFIG.DEVIN_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.HISTORY_SHEET_NAME);
    // Initialize header
    const header = ['処理日', 'ブイキューブ発注番号', 'チーム名'];
    sheet.getRange(1, 1, 1, 3).setValues([header]).setFontWeight('bold');
  }
  return sheet;
}

/**
 * Reads the master list of teams from the History Sheet (and potentially others?).
 * Since the user mentioned "History of Added Teams", this sheet might serve as the master list source.
 * It contains all previously processed teams.
 * * @returns {Array<string>} List of unique team names in the master record.
 */
function getMasterTeamList() {
  const ss = SpreadsheetApp.openById(CONFIG.MAIL_SETTINGS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.MAIL_SETTINGS_SHEET_NAME);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Header only
  
  // Assuming Team Name is in Column A (Index 0) - Check `COLUMN_DEF` usage in other files
  // In `extractDataAsMap`, it finds index by header name.
  const headers = data[0];
  const teamNameIdx = headers.indexOf(COLUMN_DEF.TEAM_NAME); // 'チーム名'
  
  if (teamNameIdx === -1) return [];
  
  const teams = data.slice(1).map(row => String(row[teamNameIdx]).trim()).filter(t => t !== '');
  return [...new Set(teams)];
}

/**
 * No longer needed as "Master List" is the Mail Settings Sheet, updated via the main tool flow.
 * But we keep the function signature to avoid breaking callers if any, though it's likely unused now.
 */
function appendTeamsToMasterList(newTeams) {
  // No-op or log warning.
  // The duplicate check logic just checks. The "Update Mail Settings" flow handles the actual addition.
}