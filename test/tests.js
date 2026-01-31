/**
 * @file tests.js
 * @brief Test suite for the Google Apps Script project.
 */

// ==========================================
// MOCK DATA AND HELPERS
// ==========================================

/**
 * Mock data for testing transformation and validation functions.
 */
const MOCK_SHEET_DATA = {
  duplicates: [
    ['チーム名', '送付先メアド(TO)', '送付先メアド(CC)'],
    ['Team A', 'a@example.com', ''],
    ['Team B', 'b@example.com', ''],
    ['Team A', 'a2@example.com', ''], // Duplicate team name
    ['Team C', 'c@example.com', ''],
    ['Team B', 'b2@example.com', ''], // Duplicate team name
    ['Team A', 'a3@example.com', '']  // Duplicate team name
  ],
  invalidEmails: [
    ['チーム名', '送付先メアド(TO)', '送付先メアド(CC)'],
    ['Team Valid', 'valid@example.com', 'cc.valid@example.com'],
    ['Team Invalid TO', 'invalid-email', 'cc@example.com'],
    ['Team Invalid CC', 'valid@example.com', 'invalid cc'],
    ['Team Invalid Both', 'bad-to@', 'bad-cc@'],
    ['Team Good TO', 'good@example.com', ''],
    ['Team Good CC', '', 'good.cc@example.com'],
    ['Team Multi TO', 'valid@example.com\ninvalid@\nvalid2@example.com', ''],
  ]
};

/**
 * A mock Sheet object for testing purposes.
 * @param {Array<Array<any>>} data The 2D array representing sheet data.
 * @returns {object} A mock sheet object with a getDataRange and getValues method.
 */
function createMockSheet(data) {
  return {
    getDataRange: () => ({
      getValues: () => data
    })
  };
}


// ==========================================
// TEST RUNNER
// ==========================================

/**
 * Main function to run all test suites.
 */
function runAllTests() {
  Logger.log('======== STARTING ALL TESTS ========');
  let failures = 0;
  const tests = [
    testUtils,
    testDataValidationSuite,
    testSourceAgainstMasterDuplicates, // Newly added test suite
    testIncrementalDuplicateCheck,
  ];

  tests.forEach(testSuite => {
    try {
      testSuite();
    } catch (e) {
      failures++;
      Logger.log(`❌ TEST SUITE FAILED: ${testSuite.name} - ${e.toString()}`);
    }
  });

  if (failures === 0) {
    Logger.log('✅✅✅ ALL TESTS PASSED SUCCESSFULLY ✅✅✅');
    SpreadsheetApp.getUi().alert('✅ All tests passed successfully!');
  } else {
    Logger.log(`❌❌❌ ${failures} TEST SUITE(S) FAILED ❌❌❌`);
    SpreadsheetApp.getUi().alert(`❌ ${failures} test suite(s) failed. Check logs for details.`);
  }
}

// ==========================================
// TEST SUITES
// ==========================================

/**
 * Test suite for functions in `utils.js`.
 */
function testUtils() {
  Logger.log('---- Running: testUtils ----');

  // Test columnLetterToIndex
  if (columnLetterToIndex('A') !== 0 || columnLetterToIndex('Z') !== 25 || columnLetterToIndex('AA') !== 26) {
    throw new Error('columnLetterToIndex failed.');
  }

  // Test columnIndexToLetter
  if (columnIndexToLetter(0) !== 'A' || columnIndexToLetter(25) !== 'Z' || columnIndexToLetter(26) !== 'AA') {
    throw new Error('columnIndexToLetter failed.');
  }

  // Test isDateLike
  if (!isDateLike(new Date()) || !isDateLike('2024-01-01') || isDateLike('not a date')) {
    throw new Error('isDateLike failed.');
  }

  // Test formatDate
  if (formatDate(new Date('2024-05-20')) !== '2024/05/20') {
     throw new Error('formatDate failed.');
  }

  // Test formatNumber
  if (formatNumber(123.456) !== 123.46) {
    throw new Error('formatNumber failed.');
  }

  Logger.log('✅ PASSED: testUtils');
}


/**
 * Test suite for functions in `dataValidation.js`.
 */
function testDataValidationSuite() {
  Logger.log('---- Running: testDataValidationSuite ----');

  // Test areEmailsValid
  if (!areEmailsValid('test@example.com\nvalid.email@domain.co.jp') || areEmailsValid('invalid-email')) {
    throw new Error('areEmailsValid failed.');
  }
  Logger.log('  -> Passed: areEmailsValid');

  // Test findDuplicates
  const mockDuplicateSheet = createMockSheet(MOCK_SHEET_DATA.duplicates);
  const duplicates = findDuplicates(mockDuplicateSheet);
  if (duplicates.length !== 2 || duplicates[0][0] !== 'Team A' || duplicates[1][0] !== 'Team B') {
    throw new Error(`findDuplicates failed. Expected 2 duplicate teams, but found ${duplicates.length}.`);
  }
  Logger.log('  -> Passed: findDuplicates');

  // Test findInvalidEmails with Map data
  const mockDataMap = new Map();
  MOCK_SHEET_DATA.invalidEmails.slice(1).forEach(row => {
    mockDataMap.set(row[0], { to: row[1], cc: row[2] });
  });

  const invalidEmails = findInvalidEmails(mockDataMap);
  const expectedInvalidCount = 4; // Team Invalid TO, Team Invalid CC, Team Invalid Both, Team Multi TO
  if (invalidEmails.length !== expectedInvalidCount) {
    throw new Error(`findInvalidEmails failed. Expected ${expectedInvalidCount} invalid teams, but found ${invalidEmails.length}.`);
  }
  const invalidTeams = invalidEmails.map(r => r[0]);
  if (!invalidTeams.includes('Team Invalid TO') || !invalidTeams.includes('Team Invalid CC') || !invalidTeams.includes('Team Invalid Both') || !invalidTeams.includes('Team Multi TO')) {
      throw new Error(`findInvalidEmails did not detect the correct invalid teams.`);
  }

  Logger.log('  -> Passed: findInvalidEmails');

  Logger.log('✅ PASSED: testDataValidationSuite');
}

/**
 * Test suite for the new cross-check duplicate logic.
 * Ensures that teams present in the source but also in the master are flagged.
 */
function testSourceAgainstMasterDuplicates() {
  Logger.log('---- Running: testSourceAgainstMasterDuplicates ----');

  // Mock Source Data (New registrations)
  // Assuming format: [TeamName, To, Cc, ...]
  const sourceData = [
    ['Team New', 'new@example.com', ''],
    ['Team Existing', 'exists@example.com', ''], // Duplicate with Master
    ['Team Another New', 'another@example.com', ''],
    ['Team Duplicate In Source', 'dup@example.com', ''] // Handled by standard duplicate check, not this one necessarily?
                                                        // Actually, checking purely against master.
  ];

  // Mock Master Data (Existing settings)
  // Assuming format from Map extract: TeamName -> Object
  const masterDataMap = new Map();
  masterDataMap.set('Team Existing', { to: 'old@example.com', cc: '' });
  masterDataMap.set('Team Old', { to: 'old2@example.com', cc: '' });

  // Call the function (not yet implemented)
  if (typeof identifyNewTeamDuplicates !== 'function') {
    throw new Error('identifyNewTeamDuplicates function is not defined.');
  }

  // Use index 0 for Team Name
  const result = identifyNewTeamDuplicates(sourceData, masterDataMap, 0);

  // Assertions
  if (!result || !Array.isArray(result)) {
    throw new Error('Result should be an array.');
  }

  // We expect "Team Existing" to be found as a duplicate.
  if (result.length !== 1) {
    throw new Error(`Expected 1 duplicate, found ${result.length}.`);
  }

  if (result[0].teamName !== 'Team Existing') {
     throw new Error(`Expected 'Team Existing' to be flagged, but got '${result[0].teamName}'.`);
  }

  Logger.log('✅ PASSED: testSourceAgainstMasterDuplicates');
}

/**
 * Test suite for the new incremental duplicate check logic.
 */
function testIncrementalDuplicateCheck() {
  Logger.log('---- Running: testIncrementalDuplicateCheck ----');

  // 1. Test checkTeamNameDuplicates
  // Update: Master Teams now include row numbers.
  const masterTeams = [
      { name: 'Team A', row: 100 },
      { name: 'Team B', row: 101 },
      { name: 'Super Team C', row: 102 }
  ];

  const newRecords = [
    { teamName: 'Team A', rowNum: 10 },        // Exact match
    { teamName: 'Team B', rowNum: 20 },        // Exact match
    { teamName: 'Team X', rowNum: 30 },        // No match
    { teamName: 'Super Team C (JP)', rowNum: 40 }, // Partial match (Master contains New? No, New contains Master)
    { teamName: 'Team', rowNum: 50 },         // Partial match (Master 'Team A' contains 'Team')
    { teamName: 'Team A', rowNum: 60 }         // Duplicate in New Records (Same as row 10)
  ];

  // Logic recap:
  // Exact: new === master
  // Partial: new.includes(master) OR master.includes(new)

  const results = checkTeamNameDuplicates(newRecords, masterTeams);

  // Expected Results:
  // 1. Team A: Exact match. Row 10 and 60. Master Row 100.
  // 2. Team B: Exact match. Row 20. Master Row 101.
  // 3. Super Team C (JP): Partial match (contains 'Super Team C'). Row 40. Master Row 102.
  // 4. Team: Partial match (contained in 'Team A', 'Team B', 'Super Team C'). Row 50. Master Rows 100, 101, 102.

  // Verify 'Team A'
  const resA = results.find(r => r.teamName === 'Team A');
  if (!resA || resA.type !== '完全' || resA.rowNums.length !== 2 || !resA.rowNums.includes(10) || !resA.rowNums.includes(60)) {
    throw new Error(`Failed check for Team A. Got: ${JSON.stringify(resA)}`);
  }
  if (String(resA.masterRow) !== '100行目') {
      // Logic update: Exact match now formats row as "X行目" too, and might use slash but here only 1 match.
      throw new Error(`Failed master row check for Team A. Expected '100行目', got ${resA.masterRow}`);
  }

  // Verify 'Team B'
  const resB = results.find(r => r.teamName === 'Team B');
  if (!resB || resB.type !== '完全' || resB.rowNums.length !== 1 || resB.rowNums[0] !== 20) {
    throw new Error(`Failed check for Team B. Got: ${JSON.stringify(resB)}`);
  }

  // Verify 'Team X' (Should not be in results)
  const resX = results.find(r => r.teamName === 'Team X');
  if (resX) {
    throw new Error(`Team X should not be in results. Got: ${JSON.stringify(resX)}`);
  }

  // Verify 'Super Team C (JP)'
  const resC = results.find(r => r.teamName === 'Super Team C (JP)');
  if (!resC || resC.type !== '部分' || !resC.matchTarget.includes('Super Team C')) {
     throw new Error(`Failed check for Super Team C (JP). Got: ${JSON.stringify(resC)}`);
  }
  // Check new metadata fields
  if (!String(resC.masterRow).includes('102')) {
       throw new Error(`Failed master row check for Super Team C (JP). Expected 102, got ${resC.masterRow}`);
  }

  // 5. Test Multiple Matches (Slash Separation)
  // 'Team' matches 'Team A' (100), 'Team B' (101), 'Super Team C' (102)
  const resTeam = results.find(r => r.teamName === 'Team');
  if (!resTeam) throw new Error('Failed to find result for "Team"');
  
  if (resTeam.type !== '部分') throw new Error('Expected "Team" to be Partial match');
  
  // Check separator
  if (!resTeam.masterRow.includes(' / ')) {
    throw new Error(`Expected slash separator in masterRow. Got: ${resTeam.masterRow}`);
  }
  // Check presence of all rows
  if (!resTeam.masterRow.includes('100行目') || !resTeam.masterRow.includes('101行目') || !resTeam.masterRow.includes('102行目')) {
    throw new Error(`Missing expected rows in multiple match. Got: ${resTeam.masterRow}`);
  }

  Logger.log('  -> Passed: checkTeamNameDuplicates logic');
  Logger.log('✅ PASSED: testIncrementalDuplicateCheck');
}