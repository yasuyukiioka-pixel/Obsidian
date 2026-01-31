/**
 * @file businessLogic.js
 * @brief This file contains the core business logic for comparing, updating, and validating data.
 */

/**
 * Compares the updated settings sheet with the latest backup to find changes.
 * This is the original, non-refactored version.
 * @deprecated This function will be replaced by compareAndUpdateResults_refactored.
 */
function compareAndUpdateResults() {
  try {
    Logger.log('Starting comparison of mail settings...');
    const ui = SpreadsheetApp.getUi();

    const mailSettingsSpreadsheet = SpreadsheetApp.openById(CONFIG.MAIL_SETTINGS_SPREADSHEET_ID);

    // --- Step 1: Find the latest backup sheet ---
    const allSheets = mailSettingsSpreadsheet.getSheets();
    const backupSheets = allSheets.filter(sheet =>
        sheet.getName().startsWith(CONFIG.MAIL_SETTINGS_SHEET_NAME) &&
        sheet.getName() !== CONFIG.MAIL_SETTINGS_SHEET_NAME
    );

    if (backupSheets.length === 0) {
      throw new Error(`バックアップシートが見つかりません。先に「2. メール設定を更新」を実行してください。`);
    }

    // Sort sheets by name descending to get the most recent one
    backupSheets.sort((a, b) => b.getName().localeCompare(a.getName()));
    const latestBackupSheet = backupSheets[0];
    Logger.log(`Comparison target backup sheet: "${latestBackupSheet.getName()}"`);

    // --- Step 2: Get sheets and data ---
    const updatedSheet = mailSettingsSpreadsheet.getSheetByName(CONFIG.MAIL_SETTINGS_SHEET_NAME);
    if (!updatedSheet) {
        throw new Error(`Sheet "${CONFIG.MAIL_SETTINGS_SHEET_NAME}" not found.`);
    }

    const updatedData = updatedSheet.getDataRange().getValues();
    const backupData = latestBackupSheet.getDataRange().getValues();

    // Dynamically find column indices from the header row
    const headers = updatedData[0];
    const teamNameIndex = headers.indexOf('チーム名');
    const toIndex = headers.indexOf('送付先メアド(TO)');
    const ccIndex = headers.indexOf('送付先メアド(CC)');

    if ([teamNameIndex, toIndex, ccIndex].includes(-1)) {
        throw new Error('Required columns (チーム名, 送付先メアド(TO), 送付先メアド(CC)) not found in the header.');
    }

    // --- Step 3: Create a map of the backup data for efficient lookup ---
    const backupMap = new Map();
    // Start from 1 to skip header
    backupData.slice(1).forEach(row => {
        const teamName = row[teamNameIndex];
        if (teamName) { // Only add rows that have a team name
            backupMap.set(teamName, {
                to: row[toIndex] || '', // Use empty string for null/undefined
                cc: row[ccIndex] || ''
            });
        }
    });

    // --- Step 4: Compare each row of the updated sheet with the backup map ---
    const differences = [];
    // Start from 1 to skip header
    updatedData.slice(1).forEach(updatedRow => {
        const teamName = updatedRow[teamNameIndex];
        if (!teamName) return; // Skip rows without a team name

        const currentTo = updatedRow[toIndex] || '';
        const currentCc = updatedRow[ccIndex] || '';

        const backupEntry = backupMap.get(teamName);

        // Condition 1-2: Not a match if backup entry doesn't exist (new team)
        // OR if To/CC emails have changed.
        if (!backupEntry || backupEntry.to !== currentTo || backupEntry.cc !== currentCc) {
            differences.push([teamName, currentTo, currentCc]);
        }
        // Condition 1-1 (match) is implicitly handled by doing nothing.
    });

    // --- Step 5: Write the differences to the results sheet ---
    let resultsSheet = mailSettingsSpreadsheet.getSheetByName(CONFIG.RESULTS_SHEET_NAME);
    if (!resultsSheet) {
        resultsSheet = mailSettingsSpreadsheet.insertSheet(CONFIG.RESULTS_SHEET_NAME);
    }

    // Clear previous results and set header
    resultsSheet.clear();
    resultsSheet.getRange("A1:C1").setValues([['チーム名', '送付先メアド(TO)の変更後', '送付先メアド(CC)の変更後']]).setFontWeight('bold');

    if (differences.length > 0) {
        // Write data starting from row 2
        resultsSheet.getRange(2, 1, differences.length, 3).setValues(differences);
        resultsSheet.autoResizeColumns(1, 3);
    }

    Logger.log(`Comparison finished. Found ${differences.length} changes.`);
    Logger.log(`比較が完了しました。\n\n新規追加または変更された項目が ${differences.length} 件見つかりました。\n詳細は「${CONFIG.RESULTS_SHEET_NAME}」シートに出力しました。`);

  } catch (error) {
    Logger.log('Error during settings comparison: ' + error.toString());
    SpreadsheetApp.getUi().alert('比較エラーが発生しました: ' + error.toString());
  }
}



/**
 * Updates the mail settings sheet based on the cleaned data list.
 * This is the original, non-refactored version.
 * @deprecated This function will be replaced by updateMailSettings_refactored.
 */
function updateMailSettings() {
  try {
    Logger.log('Starting mail settings update process...');
    const ui = SpreadsheetApp.getUi();

    // --- Step 0: Backup the target sheet ---
    const mailSettingsSpreadsheet = SpreadsheetApp.openById(CONFIG.MAIL_SETTINGS_SPREADSHEET_ID);
    const targetSheet = mailSettingsSpreadsheet.getSheetByName(CONFIG.MAIL_SETTINGS_SHEET_NAME);

    if (!targetSheet) {
      throw new Error(`Sheet "${CONFIG.MAIL_SETTINGS_SHEET_NAME}" not found in the mail settings spreadsheet.`);
    }

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmm');
    const backupSheetName = `${CONFIG.MAIL_SETTINGS_SHEET_NAME}${timestamp}`;
    targetSheet.copyTo(mailSettingsSpreadsheet).setName(backupSheetName);
    Logger.log(`Successfully created backup sheet: "${backupSheetName}"`);


    // --- Get Source and Target Data ---
    const sourceSpreadsheet = SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);
    const sourceSheet = sourceSpreadsheet.getSheetByName(CONFIG.CLEANED_DATA_SHEET_NAME);
    if (!sourceSheet) {
        throw new Error(`Cleaned data sheet "${CONFIG.CLEANED_DATA_SHEET_NAME}" not found. Please run the data transfer first.`);
    }

    const sourceData = sourceSheet.getDataRange().getValues();
    const targetData = targetSheet.getDataRange().getValues();

    // Find column indices dynamically from headers
    const sourceHeaders = sourceData[0];
    const targetHeaders = targetData[0];

    const sourceTeamNameIndex = sourceHeaders.indexOf('チーム名');
    const sourceToIndex = sourceHeaders.indexOf('稼働率レポート送付先（To）');
    const sourceCcIndex = sourceHeaders.indexOf('稼働率レポート送付先（CC）');

    const targetTeamNameIndex = targetHeaders.indexOf('チーム名');
    const targetToIndex = targetHeaders.indexOf('送付先メアド(TO)');
    const targetCcIndex = targetHeaders.indexOf('送付先メアド(CC)');

    // Validate that all required columns were found
    if ([sourceTeamNameIndex, sourceToIndex, sourceCcIndex, targetTeamNameIndex, targetToIndex, targetCcIndex].includes(-1)) {
        throw new Error('Could not find one or more required columns (チーム名, To, CC) in the source or target sheets. Please check the column headers.');
    }


    // --- Step 1: Compare, Update, or Append ---
    Logger.log('Comparing and updating/appending data...');
    const sourceTeams = sourceData.slice(1); // Exclude header row
    const targetTeamNames = targetData.map(row => row[targetTeamNameIndex]);
    let updatedCount = 0;
    let appendedCount = 0;

    sourceTeams.forEach(sourceRow => {
        const teamNameToFind = sourceRow[sourceTeamNameIndex];
        const toEmails = sourceRow[sourceToIndex];
        const ccEmails = sourceRow[sourceCcIndex];

        if (!teamNameToFind) return; // Skip if team name is empty

        // Find the row index in the target sheet
        const targetRowIndex = targetTeamNames.indexOf(teamNameToFind, 1); // Start search from index 1 to skip header

        if (targetRowIndex !== -1) {
            // --- Step 1-1: Match found, update the row ---
            targetSheet.getRange(targetRowIndex + 1, targetToIndex + 1).setValue(toEmails);
            targetSheet.getRange(targetRowIndex + 1, targetCcIndex + 1).setValue(ccEmails);
            updatedCount++;
            Logger.log(`Updated team: "${teamNameToFind}"`);
        } else {
            // --- Step 1-2: No match, append a new row ---
            const newRow = [];
            newRow[targetTeamNameIndex] = teamNameToFind;
            newRow[targetToIndex] = toEmails;
            newRow[targetCcIndex] = ccEmails;
            targetSheet.appendRow(newRow);
            appendedCount++;
            Logger.log(`Appended new team: "${teamNameToFind}"`);
        }
    });

    Logger.log('Mail settings update completed successfully.');
    Logger.log(`メール設定の更新が完了しました！\n\n更新されたチーム: ${updatedCount}件\n追加されたチーム: ${appendedCount}件\n\nバックアップシート: "${backupSheetName}" を作成しました。`);

  } catch (error) {
    Logger.log('Error during mail settings update: ' + error.toString());
  }
}

/**
 * @function compareAndUpdateResults_refactored
 * @brief Orchestrates the refactored process of comparing the updated settings sheet with a backup.
 *
 * This function performs a series of checks and data processing steps:
 * 1. Identifies the most recent backup sheet.
 * 2. Checks for duplicate team names in the current settings.
 * 3. Validates the format of email addresses in the TO and CC fields.
 * 4. Extracts data from both the current and backup sheets into structured Maps.
 * 5. Compares the two data maps to identify new, updated, or deleted entries.
 * 6. Consolidates all findings (duplicates, invalid emails, changes) into a final report.
 * 7. Writes the consolidated results to the designated report sheet.
 *
 * @throws {Error} If the main settings sheet or any backup sheets cannot be found.
 */
function compareAndUpdateResults_refactored() {
  try {
    const startTime = new Date();
    Logger.log('比較処理(リファクタリング版)を開始します...');

    const mailSettingsSpreadsheet = SpreadsheetApp.openById(CONFIG.MAIL_SETTINGS_SPREADSHEET_ID);
    const updatedSheet = mailSettingsSpreadsheet.getSheetByName(CONFIG.MAIL_SETTINGS_SHEET_NAME);
    if (!updatedSheet) throw new Error(`シート "${CONFIG.MAIL_SETTINGS_SHEET_NAME}" が見つかりません。`);

    const latestBackupSheet = getLatestBackupSheet(mailSettingsSpreadsheet);

    // データ抽出
    const updatedDataMap = extractDataAsMap(updatedSheet);
    const backupDataMap = extractDataAsMap(latestBackupSheet);

    // 1. 重複チェック
    const duplicateErrors = findDuplicates(updatedSheet);

    // 2. メールアドレス形式チェック
    const emailFormatErrors = findInvalidEmails(updatedDataMap);


    // 4. 差分抽出
    const changes = findChanges(updatedDataMap, backupDataMap);

    // 5. 結果結合
    const finalResults = [...duplicateErrors, ...emailFormatErrors, ...changes];

    // 書き込み
    writeResults(mailSettingsSpreadsheet, finalResults);

    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    Logger.log(`比較処理完了: 処理時間 ${duration}秒 (出力件数: ${finalResults.length}件)`);

  } catch (error) {
    Logger.log('比較処理中にエラーが発生しました: ' + error.toString());
    SpreadsheetApp.getUi().alert('比較処理中にエラーが発生しました: ' + error.toString());
  }
}

/**
 * @function checkForDuplicatesAndNotify
 * @brief Checks if any new teams in the source data already exist in the master settings.
 * If duplicates are found, it sends an email notification.
 * Uses the raw source data (Procurement Input) and mapping rules.
 * * @param {boolean} isInteractive - If true, displays UI alerts (for manual execution).
 */
function checkForDuplicatesAndNotify(isInteractive = false) {
  try {
    Logger.log('新規登録の重複チェックを開始します...');

    // 1. Get Raw Source Data (Procurement Input Sheet)
    const procurementSpreadsheet = SpreadsheetApp.openById(CONFIG.PROCUREMENT_SPREADSHEET_ID);
    const procurementSheet = procurementSpreadsheet.getSheetByName(CONFIG.PROCUREMENT_SHEET_NAME);
    
    if (!procurementSheet) {
      throw new Error(`調達側シート "${CONFIG.PROCUREMENT_SHEET_NAME}" が見つかりません。config.jsのPROCUREMENT_SHEET_NAMEを確認してください。`);
    }
    
    // Read the data range
    Logger.log(`読み込み対象範囲: ${CONFIG.PROCUREMENT_DATA_RANGE}`);
    const range = procurementSheet.getRange(CONFIG.PROCUREMENT_DATA_RANGE);
    const sourceData = range.getValues();
    
    if (!sourceData || sourceData.length === 0) {
       Logger.log('ソースデータがありません（0件）。');
       if (isInteractive) SpreadsheetApp.getUi().alert('チェック対象のソースデータがありません。');
       return;
    }
    Logger.log(`ソースデータを ${sourceData.length} 行読み込みました。`);

    // 2. Determine Source Column Index for "Team Name" by Header Search
    // Assuming Header is at CONFIG.PROCUREMENT_HEADER_ROW_NUM (Row 18)
    const headerRowIdx = CONFIG.PROCUREMENT_HEADER_ROW_NUM || 18;
    
    // Read the entire header row to find 'チーム名'
    // getRange(row, column, numRows, numColumns) -> (18, 1, 1, LastColumn)
    const lastCol = procurementSheet.getLastColumn();
    const headerValues = procurementSheet.getRange(headerRowIdx, 1, 1, lastCol).getValues()[0];
    
    let sourceTeamNameIndex = -1;
    
    // Find the index of "チーム名" or "Team Name"
    const headerIndex = headerValues.findIndex(h => String(h).includes('チーム名') || String(h).includes('Team Name'));
    
    if (headerIndex !== -1) {
        // headerIndex is 0-based index in the sheet (0 = Column A, 1 = Column B...)
        
        // We need the index relative to CONFIG.PROCUREMENT_DATA_RANGE.
        // If data range is 'B21:L', it starts at Column B (Index 1).
        // relativeIndex = headerIndex - startIndex
        
        const rangeStartColumnLetter = CONFIG.PROCUREMENT_DATA_RANGE.split(':')[0].replace(/[0-9]/g, '');
        const rangeStartColumnIndex = columnLetterToIndex(rangeStartColumnLetter); // B -> 1
        
        sourceTeamNameIndex = headerIndex - rangeStartColumnIndex;
        
        Logger.log(`ヘッダー行(${headerRowIdx}行目)から 'チーム名' を検出しました。`);
        Logger.log(`絶対列インデックス: ${headerIndex} (列 ${columnIndexToLetter(headerIndex)})`);
        Logger.log(`データ範囲開始列: ${rangeStartColumnIndex} (列 ${rangeStartColumnLetter})`);
        Logger.log(`相対列インデックス: ${sourceTeamNameIndex}`);
        
    } else {
       Logger.log(`ヘッダー行(${headerRowIdx}行目)に 'チーム名' が見つかりませんでした。`);
       if (isInteractive) SpreadsheetApp.getUi().alert(`ヘッダー行(${headerRowIdx}行目)に 'チーム名' 列が見つかりません。`);
       return;
    }
    
    if (sourceTeamNameIndex < 0) {
       Logger.log(`計算された列インデックスが無効です (${sourceTeamNameIndex})。データ範囲の設定を確認してください。`);
       if (isInteractive) SpreadsheetApp.getUi().alert(`列インデックス計算エラー (${sourceTeamNameIndex})。`);
       return;
    }
    
    // Log sample values from the identified column
    Logger.log('--- チーム名列のサンプルデータ (最初の5件) ---');
    for (let i = 0; i < Math.min(5, sourceData.length); i++) {
        // Verify index is within bounds
        if (sourceData[i].length > sourceTeamNameIndex) {
            Logger.log(`Row ${i+1}: ${sourceData[i][sourceTeamNameIndex]}`);
        } else {
            Logger.log(`Row ${i+1}: [Out of bounds]`);
        }
    }
    Logger.log('-------------------------------------------');


    // 3. Get Master Data
    const mailSettingsSpreadsheet = SpreadsheetApp.openById(CONFIG.MAIL_SETTINGS_SPREADSHEET_ID);
    const masterSheet = mailSettingsSpreadsheet.getSheetByName(CONFIG.MAIL_SETTINGS_SHEET_NAME);
    if (!masterSheet) {
      throw new Error(`マスタシート "${CONFIG.MAIL_SETTINGS_SHEET_NAME}" が見つかりません。`);
    }
    const masterDataMap = extractDataAsMap(masterSheet);
    Logger.log(`マスタデータから ${masterDataMap.size} 件のチームを読み込みました。`);

    // 4. Identify Duplicates
    const duplicates = identifyNewTeamDuplicates(sourceData, masterDataMap, sourceTeamNameIndex);

    if (duplicates.length === 0) {
      Logger.log('重複は見つかりませんでした。');
      if (isInteractive) SpreadsheetApp.getUi().alert('重複するチームは見つかりませんでした。');
      return;
    }

    // 5. Notify
    Logger.log(`${duplicates.length} 件の重複が見つかりました。通知を送信します。`);
    
    const recipient = CONFIG.NOTIFICATION_EMAIL;
    if (!recipient) {
      Logger.log('通知先メールアドレスが設定されていません。');
      if (isInteractive) SpreadsheetApp.getUi().alert('通知先メールアドレスが設定されていません。config.jsを確認してください。');
      return;
    }

    const subject = `【Jules】新規登録チームの重複警告 (${duplicates.length}件)`;
    let body = '以下のチームは既にマスタ設定に存在します。場所違い（例：神田と浅草）などの可能性がありますので確認してください。\n\n';
    
    // Calculate source row number
    const startRow = parseInt(CONFIG.PROCUREMENT_DATA_RANGE.match(/[0-9]+/)[0], 10);

    duplicates.forEach(d => {
      // d.sourceRowIndex is 1-based index in the array.
      // Actual Sheet Row = StartRow + (Index - 1)
      const actualRow = startRow + (d.sourceRowIndex - 1);
      body += `- ${d.teamName} (行番号: ${actualRow})\n`;
    });

    body += '\n確認をお願いします。';

    GmailApp.sendEmail(recipient, subject, body);
    Logger.log(`通知メールを ${recipient} に送信しました。`);

    if (isInteractive) {
      SpreadsheetApp.getUi().alert(`${duplicates.length} 件の重複が見つかりました。\n担当者 (${recipient}) にメール通知を送信しました。`);
    }

  } catch (error) {
    Logger.log('重複チェック処理中にエラーが発生しました: ' + error.toString());
    if (isInteractive) SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.toString());
  }
}

/**
 * @function extractDataAsMap
 * @brief Extracts and transforms data from a sheet into a Map for easy lookup.
 *
 * The function reads the display values from the sheet to handle formatted dates/numbers correctly.
 * It normalizes all text data by trimming whitespace and removing invisible characters to ensure
 * consistent comparisons. Each row is stored in a Map with the team name as the key.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to extract data from.
 * @returns {Map<string, object>} A map where keys are normalized team names and values are objects
 * containing the row data.
 * @throws {Error} If essential columns (Team Name, TO, CC) are missing.
 */
function extractDataAsMap(sheet) {
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length === 0) return new Map();

  const headers = data[0];
  const indices = {
    teamName: headers.indexOf(COLUMN_DEF.TEAM_NAME),
    to: headers.indexOf(COLUMN_DEF.TO),
    cc: headers.indexOf(COLUMN_DEF.CC),
    period: headers.indexOf(COLUMN_DEF.PERIOD),
    startTime: headers.indexOf(COLUMN_DEF.START_TIME),
    endTime: headers.indexOf(COLUMN_DEF.END_TIME),
    holiday: headers.indexOf(COLUMN_DEF.HOLIDAY)
  };

  if (indices.teamName === -1 || indices.to === -1 || indices.cc === -1) {
    throw new Error(`シート「${sheet.getName()}」に必須列が見つかりません。`);
  }

  const dataMap = new Map();

  const normalize = (str) => String(str)
    .normalize('NFKC')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .trim();

  const checkCleaning = (raw, cleaned) => {
    if (!raw) return false;
    return raw.replace(/\s/g, '') !== cleaned.replace(/\s/g, '');
  };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rawTeamName = row[indices.teamName];

    if (rawTeamName) {
      const normalizedTeamName = normalize(rawTeamName);

      const rawTo = indices.to !== -1 ? String(row[indices.to]) : '';
      const rawCc = indices.cc !== -1 ? String(row[indices.cc]) : '';

      const cleanTo = normalize(rawTo);
      const cleanCc = normalize(rawCc);

      const cleanedColumns = [];
      if (checkCleaning(rawTo, cleanTo)) cleanedColumns.push('TO');
      if (checkCleaning(rawCc, cleanCc)) cleanedColumns.push('CC');

      const entry = {
        to: cleanTo,
        cc: cleanCc,
        period: indices.period !== -1 ? normalize(row[indices.period]) : '',
        startTime: indices.startTime !== -1 ? normalize(row[indices.startTime]) : '',
        endTime: indices.endTime !== -1 ? normalize(row[indices.endTime]) : '',
        holiday: indices.holiday !== -1 ? normalize(row[indices.holiday]) : '',
        cleanedCols: cleanedColumns
      };

      dataMap.set(normalizedTeamName, entry);
    }
  }
  return dataMap;
}

/**
 * @function findChanges
 * @brief Compares the data from the updated sheet and the backup sheet to identify differences.
 *
 * It iterates through the updated data and categorizes each team as:
 * - 'C' (Created): The team exists in the new data but not in the backup.
 * - 'U' (Updated): The team exists in both, but some of its data has changed.
 * - 'D' (Deleted): The team exists in the backup but not in the new data.
 * It also generates remarks for any observed changes, such as modifications to specific columns
 * or automated cleaning of invisible characters.
 *
 * @param {Map<string, object>} updatedMap The map of data from the current settings sheet.
 * @param {Map<string, object>} backupMap The map of data from the backup sheet.
 * @returns {Array<Array<string>>} A 2D array representing the rows to be written to the results sheet.
 */
function findChanges(updatedMap, backupMap) {
  const differences = [];
  const processedTeams = new Set();

  const HEADER_LABEL_MAP = {
    period: COLUMN_DEF.PERIOD,
    startTime: COLUMN_DEF.START_TIME,
    endTime: COLUMN_DEF.END_TIME,
    holiday: COLUMN_DEF.HOLIDAY
  };

  for (const [teamName, updatedEntry] of updatedMap.entries()) {
    const backupEntry = backupMap.get(teamName);

    let remarksList = [];

    if (updatedEntry.cleanedCols && updatedEntry.cleanedCols.length > 0) {
      remarksList.push(`※不可視文字を自動削除(${updatedEntry.cleanedCols.join(', ')})`);
    }

    if (!backupEntry) {
      differences.push([
        teamName,
        updatedEntry.to,
        updatedEntry.cc,
        'C',
        remarksList.join('、')
      ]);
    } else {
      let isChanged = false;

      if (updatedEntry.to !== backupEntry.to || updatedEntry.cc !== backupEntry.cc) {
        isChanged = true;
      }

      Object.keys(HEADER_LABEL_MAP).forEach(key => {
        if (updatedEntry[key] !== backupEntry[key]) {
          isChanged = true;
          remarksList.push(`列：${HEADER_LABEL_MAP[key]}`);
        }
      });

      if (isChanged) {
        const isMailEmpty = (updatedEntry.to === '' && updatedEntry.cc === '');
        const isSettingChanged = Object.keys(HEADER_LABEL_MAP).some(k => updatedEntry[k] !== backupEntry[k]);
        const changeType = (isMailEmpty && !isSettingChanged) ? 'D' : 'U';

        differences.push([
          teamName,
          updatedEntry.to,
          updatedEntry.cc,
          changeType,
          remarksList.join('\n')
        ]);
      }
    }
    processedTeams.add(teamName);
  }

  for (const [teamName, backupEntry] of backupMap.entries()) {
    if (!processedTeams.has(teamName)) {
      differences.push([teamName, '', '', 'D', '']);
    }
  }

  return differences;
}


/**
 * @function updateMailSettings_refactored
 * @brief Orchestrates the process of updating the main mail settings sheet from a source of truth.
 *
 * This function follows a safe update procedure:
 * 1. Backs up the current mail settings sheet.
 * 2. Fetches the clean, definitive data from the source sheet (`CLEANED_DATA_SHEET_NAME`).
 * 3. Reconstructs the final dataset in memory by comparing the source and target data,
 * identifying what needs to be added, updated, or deleted.
 * 4. Wipes the old data from the target sheet (preserving headers) and writes the new,
 * reconstructed data.
 *
 * @throws {Error} If the source or target sheets cannot be found.
 */
function updateMailSettings_refactored() {
  try {
    Logger.log('メール設定の更新処理を開始します...');

    const mailSettingsSpreadsheet = SpreadsheetApp.openById(CONFIG.MAIL_SETTINGS_SPREADSHEET_ID);
    const targetSheet = mailSettingsSpreadsheet.getSheetByName(CONFIG.MAIL_SETTINGS_SHEET_NAME);
    if (!targetSheet) {
      throw new Error(`シート "${CONFIG.MAIL_SETTINGS_SHEET_NAME}" が見つかりません。`);
    }

    const sourceData = getSheetData(CONFIG.TARGET_SPREADSHEET_ID, CONFIG.CLEANED_DATA_SHEET_NAME);
    const targetData = targetSheet.getDataRange().getValues();

    const backupSheetName = backupSheet(targetSheet);

    const result = reconstructFinalData(sourceData, targetData);

    updateSheetWithData(targetSheet, result.finalData);

    Logger.log(`メール設定の更新が完了しました！
更新: ${result.stats.updated}件, 追加: ${result.stats.appended}件, 削除: ${result.stats.deleted}件
バックアップシート: "${backupSheetName}" を作成しました。`);

  } catch (error) {
    Logger.log('メール設定の更新中にエラーが発生しました: ' + error.toString());
  }
}

/**
 * @function reconstructFinalData
 * @brief Rebuilds the dataset for the mail settings sheet based on a source of truth.
 *
 * It compares the `sourceData` (considered the correct state) with the `targetData` (the current state).
 * - If a team is in the source and target, it's updated if necessary.
 * - If a team is in the source but not the target, it's added.
 * - If a team is in the target but not the source, it's removed.
 * This ensures the target sheet is perfectly synchronized with the source.
 *
 * @param {Array<Array<any>>} sourceData The master data from the "cleaned" sheet.
 * @param {Array<Array<any>>} targetData The current data from the "mail settings" sheet.
 * @returns {{finalData: Array<Array<any>>, stats: {updated: number, appended: number, deleted: number}}}
 * An object containing the final, reconstructed data and statistics about the changes.
 */
function reconstructFinalData(sourceData, targetData) {
  const sourceHeaders = sourceData[0];
  const targetHeaders = targetData[0];
  const sourceTeamIdx = sourceHeaders.indexOf('チーム名');
  const sourceToIdx = sourceHeaders.indexOf('稼働率レポート送付先（To）');
  const sourceCcIdx = sourceHeaders.indexOf('稼働率レポート送付先（CC）');
  const targetTeamIdx = targetHeaders.indexOf('チーム名');
  const targetToIdx = targetHeaders.indexOf('送付先メアド(TO)');
  const targetCcIdx = targetHeaders.indexOf('送付先メアド(CC)');

  const sourceMap = new Map(sourceData.slice(1).map(row => [row[sourceTeamIdx], { to: row[sourceToIdx], cc: row[sourceCcIdx] }]));
  const targetMap = new Map(targetData.slice(2).map(row => [row[targetTeamIdx], row]));

  const finalData = [];
  const stats = { updated: 0, appended: 0, deleted: 0 };

  for (const [teamName, originalRow] of targetMap.entries()) {
    if (sourceMap.has(teamName)) {
      const sourceInfo = sourceMap.get(teamName);
      const updatedRow = [...originalRow];
      if (updatedRow[targetToIdx] !== sourceInfo.to || updatedRow[targetCcIdx] !== sourceInfo.cc) {
        updatedRow[targetToIdx] = sourceInfo.to;
        updatedRow[targetCcIdx] = sourceInfo.cc;
        stats.updated++;
      }
      finalData.push(updatedRow);
    } else {
      stats.deleted++;
    }
  }

  for (const [teamName, sourceInfo] of sourceMap.entries()) {
    if (!targetMap.has(teamName)) {
      const newRow = new Array(targetHeaders.length).fill('');
      newRow[targetTeamIdx] = teamName;
      newRow[targetToIdx] = sourceInfo.to;
      newRow[targetCcIdx] = sourceInfo.cc;
      finalData.push(newRow);
      stats.appended++;
    }
  }

  return { finalData, stats };
}

/**
 * Retrieves new records from the Procurement Sheet that appear after the last processed Order Number.
 * * @returns {object} An object containing:
 * - newRecords: Array of row objects (with 'teamName', 'orderNumber', 'rowNum')
 * - lastOrderNumber: The last order number found in the new records (to update state later)
 */
function getNewProcurementRecords() {
  const lastProcessedOrderNum = getLastProcessedOrderNumber();
  Logger.log(`前回の発注番号: ${lastProcessedOrderNum}`);

  const procurementSpreadsheet = SpreadsheetApp.openById(CONFIG.PROCUREMENT_SPREADSHEET_ID);
  const procurementSheet = procurementSpreadsheet.getSheetByName(CONFIG.PROCUREMENT_SHEET_NAME);
  if (!procurementSheet) {
    throw new Error(`シート "${CONFIG.PROCUREMENT_SHEET_NAME}" が見つかりません。`);
  }

  // Read header to find columns
  const headerRowIdx = CONFIG.PROCUREMENT_HEADER_ROW_NUM || 18;
  const lastCol = procurementSheet.getLastColumn();
  const headerValues = procurementSheet.getRange(headerRowIdx, 1, 1, lastCol).getValues()[0];

  // Find column indices
  const teamNameIdx = headerValues.findIndex(h => String(h).includes(COLUMN_DEF.TEAM_NAME) || String(h).includes('Team Name'));
  const orderNumIdx = headerValues.findIndex(h => String(h).includes(COLUMN_DEF.ORDER_NUMBER));

  if (teamNameIdx === -1) throw new Error(`ヘッダーに "${COLUMN_DEF.TEAM_NAME}" が見つかりません。`);
  if (orderNumIdx === -1) throw new Error(`ヘッダーに "${COLUMN_DEF.ORDER_NUMBER}" が見つかりません。`);

  // Read data
  // Assuming data starts 3 rows after header (based on previous config B21 vs Header 18)
  // Let's use the explicit data range start if available, or assume Header + 3.
  const dataStartRow = parseInt(CONFIG.PROCUREMENT_DATA_RANGE.match(/[0-9]+/)[0], 10);
  const lastRow = procurementSheet.getLastRow();

  if (lastRow < dataStartRow) {
    Logger.log("データ行が存在しません。");
    return { newRecords: [], lastOrderNumber: lastProcessedOrderNum };
  }

  const dataRange = procurementSheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastCol);
  const dataValues = dataRange.getValues();

  const newRecords = [];
  let foundLastProcessed = false;
  let lastOrderNumber = lastProcessedOrderNum;
  let lastTeamName = '';

  // Logic:
  // Iterate through rows.
  // If we haven't found the "Last Processed ID" yet, we skip (unless it's empty, then we take all?)
  // Actually, if lastProcessedOrderNum is empty, we take ALL records.
  // If it's not empty, we look for the row matching it. All rows AFTER that match are "new".
  // Note: This assumes the list is ordered by time/ID.

  if (!lastProcessedOrderNum) {
    // Take all records
    foundLastProcessed = true; 
    Logger.log("前回の発注番号がないため、全レコードを対象とします。");
  }

  let startCollecting = !lastProcessedOrderNum; // If no last ID, start immediately

  for (let i = 0; i < dataValues.length; i++) {
    const row = dataValues[i];
    const teamName = row[teamNameIdx];
    const orderNum = String(row[orderNumIdx]).trim();

    // Skip empty rows
    if (!teamName && !orderNum) continue;

    if (!startCollecting) {
      if (orderNum === lastProcessedOrderNum) {
        startCollecting = true;
        Logger.log(`前回の発注番号 (${orderNum}) を発見しました。次の行から収集を開始します。`);
      }
      continue;
    }

    // Collecting new records
    if (teamName) { // Ensure team name exists
      const trimmedTeamName = String(teamName).trim();
      newRecords.push({
        teamName: trimmedTeamName,
        orderNumber: orderNum,
        rowNum: dataStartRow + i
      });
      // Keep track of the latest order number seen
      if (orderNum) {
          lastOrderNumber = orderNum;
          lastTeamName = trimmedTeamName;
      }
    }
  }

  Logger.log(`新規レコード数: ${newRecords.length}`);
  return { newRecords, lastOrderNumber, lastTeamName };
}

/**
 * Main function to execute the incremental duplicate check and generate a report.
 * Triggered from the menu.
 */
function runIncrementalDuplicateCheck() {
  try {
    const ui = SpreadsheetApp.getUi();
    Logger.log('前回以降の重複チェックを開始します...');

    // 1. Get Master Team List
    // Modified to return objects { name, row } instead of just strings
    const masterTeams = getMasterTeamList();
    Logger.log(`マスタ登録済みチーム数: ${masterTeams.length}`);

    // 2. Get New Records from Procurement Sheet
    const { newRecords, lastOrderNumber, lastTeamName } = getNewProcurementRecords();
    
    if (newRecords.length === 0) {
      Logger.log('新規レコードはありません。');
      ui.alert('前回チェック以降の新規レコードはありませんでした。');
      return;
    }

    // 3. Perform Check
    const duplicates = checkTeamNameDuplicates(newRecords, masterTeams);
    Logger.log(`重複検出数: ${duplicates.length}`);

    // 4. Generate Report
    let message = '';
    if (duplicates.length > 0) {
      outputDuplicateReport(duplicates);
      message = `重複チェック完了: ${duplicates.length} 件の重複が見つかりました。\n詳細は「${CONFIG.DUPLICATE_REPORT_SHEET_NAME}」シートを確認してください。`;
    } else {
      message = `重複チェック完了: 新規レコード ${newRecords.length} 件中に重複は見つかりませんでした。`;
    }

    // 5. Update State (Last Processed Order Number)
    // Update automatically as per requirement.
    if (lastOrderNumber && lastTeamName) {
      updateLastProcessedOrderNumber(lastOrderNumber, lastTeamName);
      message += `\n\n処理履歴を更新しました (最終発注番号: ${lastOrderNumber})`;
    }
    
    // ui.alert(message);

  } catch (error) {
    Logger.log('エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.toString());
  }
}

/**
 * Outputs the duplicate report to a spreadsheet sheet.
 * @param {Array<object>} duplicates - Result objects from checkTeamNameDuplicates
 */
function outputDuplicateReport(duplicates) {
  const ss = SpreadsheetApp.openById(CONFIG.DEVIN_SPREADSHEET_ID); // Outputting to Devin Sheet's file? Or Procurement?
  // User didn't specify file, but "Duplicate Report Sheet" name is in config.
  // Let's use the Mail Settings / Devin Spreadsheet as the workspace.
  
  let sheet = ss.getSheetByName(CONFIG.DUPLICATE_REPORT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.DUPLICATE_REPORT_SHEET_NAME);
  }
  
  sheet.clear();
  
  // Header
  // Format: 種別, チーム名, 行番号, 重複対象(参考), マスタ合致行, 合致キーワード
  const header = [['種別', 'チーム名', '行番号', '重複対象(参考)', 'マスタ合致行', '合致キーワード']];
  sheet.getRange(1, 1, 1, 6).setValues(header).setFontWeight('bold');
  
  // Data
  const rows = duplicates.map(d => [
    d.type,
    d.teamName,
    d.rowNums.join(' : '), // User requested format "10:50:70"
    d.matchTarget,
    d.masterRow,
    d.matchedKeyword
  ]);
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    sheet.autoResizeColumns(1, 6);
  }
}

/**
 * Checks for duplicates between new team names and the master team list.
 * Supports both Exact Match and Partial Match.
 * * @param {Array<object>} newRecords - Array of objects { teamName, rowNum, ... }
 * @param {Array<object>} masterTeams - Array of objects { name: string, row: number } from the master list.
 * @returns {Array<object>} Array of result objects { type: '完全'|'部分', teamName, rowNums: [], matchTarget: string, masterRow: string, matchedKeyword: string }
 */
function checkTeamNameDuplicates(newRecords, masterTeams) {
  const results = [];
  
  // Normalize Helper
  const norm = str => String(str).trim();

  newRecords.forEach(record => {
    const newTeam = norm(record.teamName);
    const rowNum = record.rowNum;

    let exactMatch = null;
    let partialMatches = [];

    // Iterate master teams
    let exactMatches = [];
    
    for (const master of masterTeams) {
      const masterName = norm(master.name);
      
      if (newTeam === masterName) {
        exactMatches.push(master);
      } else if (newTeam.includes(masterName) || masterName.includes(newTeam)) {
        partialMatches.push(master);
      }
    }

    if (exactMatches.length > 0) {
      // Handle multiple exact matches (if any, though rare)
      const targetNames = exactMatches.map(m => m.name).join(' / ');
      const targetRows = exactMatches.map(m => `${m.row}行目`).join(' / ');
      const keywords = exactMatches.map(m => m.name).join(' / ');
      
      addResult(results, '完全', newTeam, rowNum, targetNames, targetRows, keywords);
      
    } else if (partialMatches.length > 0) {
      // Collect all partial matches
      const targetNames = partialMatches.map(m => m.name).join(' / ');
      const targetRows = partialMatches.map(m => `${m.row}行目`).join(' / ');
      const keywords = partialMatches.map(m => m.name).join(' / '); // Keyword is the master team name that matched

      addResult(results, '部分', newTeam, rowNum, targetNames, targetRows, keywords);
    }
  });

  return results;
}

/**
 * Helper to add or update results in the list.
 * Groups by (Type + TeamName).
 */
function addResult(results, type, teamName, rowNum, matchTarget, masterRow, matchedKeyword) {
  const existing = results.find(r => r.type === type && r.teamName === teamName);
  
  if (existing) {
    if (!existing.rowNums.includes(rowNum)) {
      existing.rowNums.push(rowNum);
    }
    // Update metadata if new matches found (simple concatenation for now to avoid losing info)
    if (!existing.matchTarget.includes(matchTarget)) {
       existing.matchTarget += ` / ${matchTarget}`;
    }
    if (!existing.masterRow.includes(masterRow)) {
       existing.masterRow += ` / ${masterRow}`;
    }
    if (!existing.matchedKeyword.includes(matchedKeyword)) {
       existing.matchedKeyword += ` / ${matchedKeyword}`;
    }
  } else {
    results.push({
      type: type,
      teamName: teamName,
      rowNums: [rowNum],
      matchTarget: matchTarget,
      masterRow: String(masterRow),
      matchedKeyword: matchedKeyword
    });
  }
}
