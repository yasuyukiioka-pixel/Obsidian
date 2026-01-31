/**
 * @file spreadsheetUtils.js
 * @brief This file contains utility functions for interacting with Google Spreadsheets.
 */

/**
 * Read headers from the source spreadsheet
 * @return {Array} 2D array of header data
 */
function readSourceHeaders() {
  try {
    const sourceSpreadsheet = SpreadsheetApp.openById(CONFIG.SOURCE_SPREADSHEET_ID);
    const sourceSheet = sourceSpreadsheet.getSheetByName(CONFIG.SOURCE_SHEET_NAME);

    if (!sourceSheet) {
      throw new Error(`Source sheet "${CONFIG.SOURCE_SHEET_NAME}" not found`);
    }

    const range = sourceSheet.getRange(CONFIG.SOURCE_TITLE_RANGE);
    const values = range.getValues();

    Logger.log(`Read ${values.length} header rows from source spreadsheet`);
    return values;

  } catch (error) {
    throw new Error('Failed to read source headers: ' + error.toString());
  }
}

/**
 * Read data from the source spreadsheet (excluding headers)
 * @return {Array} 2D array of source data
 */
function readSourceData() {
  try {
    const sourceSpreadsheet = SpreadsheetApp.openById(CONFIG.SOURCE_SPREADSHEET_ID);
    const sourceSheet = sourceSpreadsheet.getSheetByName(CONFIG.SOURCE_SHEET_NAME);

    if (!sourceSheet) {
      throw new Error(`Source sheet "${CONFIG.SOURCE_SHEET_NAME}" not found`);
    }

    const range = sourceSheet.getRange(CONFIG.SOURCE_DATA_RANGE);
    const values = range.getValues();

    // Filter out empty rows
    const filteredData = values.filter(row => row.some(cell => cell !== ''));

    Logger.log(`Read ${filteredData.length} data rows from source spreadsheet`);
    return filteredData;

  } catch (error) {
    throw new Error('Failed to read source data: ' + error.toString());
  }
}

/**
 * Read column mapping rules from the mapping spreadsheet
 * @return {Array} Array of mapping objects with sourceIndex, targetIndex, sourceColumn, targetColumn
 */
function readColumnMappingRules() {
  try {
    const mappingSpreadsheet = SpreadsheetApp.openById(CONFIG.MAPPING_SPREADSHEET_ID);
    const mappingSheet = mappingSpreadsheet.getSheetByName(CONFIG.MAPPING_SHEET_NAME);

    if (!mappingSheet) {
      throw new Error(`Mapping sheet "${CONFIG.MAPPING_SHEET_NAME}" not found`);
    }

    const range = mappingSheet.getRange(CONFIG.MAPPING_DATA_RANGE);
    const values = range.getValues();

    const mappingRules = [];

    // Process each mapping rule row
    values.forEach((row, index) => {
      const sourceColumn = row[0] ? row[0].toString().trim().toUpperCase() : '';
      const targetColumn = row[1] ? row[1].toString().trim().toUpperCase() : '';
      const description = row[2] ? row[2].toString().trim() : '';

      // Skip empty rows
      if (!sourceColumn || !targetColumn) {
        return;
      }

      // Convert column letters to indices (A=0, B=1, etc.)
      const sourceIndex = columnLetterToIndex(sourceColumn);
      const targetIndex = columnLetterToIndex(targetColumn);

      if (sourceIndex === -1 || targetIndex === -1) {
        Logger.log(`Warning: Invalid column mapping at row ${index + 2}: ${sourceColumn} → ${targetColumn}`);
        return;
      }

      mappingRules.push({
        sourceColumn: sourceColumn,
        targetColumn: targetColumn,
        sourceIndex: sourceIndex,
        targetIndex: targetIndex,
        description: description
      });
    });

    Logger.log(`Read ${mappingRules.length} column mapping rules from spreadsheet`);
    return mappingRules;

  } catch (error) {
    throw new Error('Failed to read column mapping rules: ' + error.toString());
  }
}

/**
 * Write data to a specific target sheet (supports 4-stage output)
 * @param {Array} headers - 2D array of header data to write
 * @param {Array} data - 2D array of data to write
 * @param {string} sheetName - Name of the target sheet to write to
 */
function writeToTargetSheet(headers, data, sheetName) {
  try {
    const targetSpreadsheet = SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);
    let targetSheet = targetSpreadsheet.getSheetByName(sheetName);

    if (!targetSheet) {
      targetSheet = targetSpreadsheet.insertSheet(sheetName);
      Logger.log(`Created new sheet: ${sheetName}`);
    }

    targetSheet.clear();
    Logger.log(`Cleared existing data from sheet: ${sheetName}`);

    writeHeadersToTargetSpreadsheet(targetSheet, headers);
    writeDataToTargetSpreadsheet(targetSheet, data);

    Logger.log(`Successfully wrote data to sheet: ${sheetName}`);

  } catch (error) {
    throw new Error(`Failed to write to sheet "${sheetName}": ` + error.toString());
  }
}

/**
 * Write data to the target spreadsheet (legacy function for backward compatibility)
 * @param {Array} headers - 2D array of header data to write
 * @param {Array} data - 2D array of data to write
 */
function writeToTargetSpreadsheet(headers, data) {
  writeToTargetSheet(headers, data, CONFIG.TARGET_SHEET_NAME);
}

/**
 * Write headers to the target spreadsheet
 * @param {Sheet} targetSheet - Target sheet object
 * @param {Array} headers - 2D array of header data to write
 */
function writeHeadersToTargetSpreadsheet(targetSheet, headers) {
  try {
    if (CONFIG.INCLUDE_HEADERS && headers && headers.length > 0) {
      const headerRange = targetSheet.getRange(CONFIG.TARGET_TITLE_RANGE);

      const headerRows = headers.length;
      const headerCols = headers[0] ? headers[0].length : 0;

      if (headerCols > 0) {
        const adjustedHeaderRange = targetSheet.getRange(
          headerRange.getRow(),
          headerRange.getColumn(),
          headerRows,
          headerCols
        );
        adjustedHeaderRange.setValues(headers);
        Logger.log(`Successfully wrote ${headerRows} header rows to target spreadsheet`);
      } else {
        Logger.log('No header columns to write');
      }
    } else {
      Logger.log('Headers not included or no header data provided');
    }
  } catch (error) {
    throw new Error('Failed to write headers to target spreadsheet: ' + error.toString());
  }
}

/**
 * Write data to the target spreadsheet
 * @param {Sheet} targetSheet - Target sheet object
 * @param {Array} data - 2D array of data to write
 */
function writeDataToTargetSpreadsheet(targetSheet, data) {
  try {
    if (data && data.length > 0) {
      const dataStartRange = targetSheet.getRange(CONFIG.TARGET_DATA_START_CELL);
      const dataRows = data.length;
      const dataCols = data[0] ? data[0].length : 0;

      if (dataCols > 0) {
        const dataRange = targetSheet.getRange(
          dataStartRange.getRow(),
          dataStartRange.getColumn(),
          dataRows,
          dataCols
        );
        dataRange.setValues(data);
        Logger.log(`Successfully wrote ${dataRows} data rows to target spreadsheet`);
      } else {
        Logger.log('No data columns to write');
      }
    } else {
      Logger.log('No data to write to target spreadsheet');
    }
  } catch (error) {
    throw new Error('Failed to write data to target spreadsheet: ' + error.toString());
  }
}

/**
 * Latest backup sheetを取得します。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - 対象のスプレッドシート。
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} - 最新のバックアップシートオブジェクト。
 */
function getLatestBackupSheet(spreadsheet) {
  const allSheets = spreadsheet.getSheets();
  const backupSheets = allSheets.filter(sheet =>
    sheet.getName().startsWith(CONFIG.MAIL_SETTINGS_SHEET_NAME) &&
    sheet.getName() !== CONFIG.MAIL_SETTINGS_SHEET_NAME
  );

  if (backupSheets.length === 0) {
    throw new Error('バックアップシートが見つかりません。先に「リストを作成」を実行してください。');
  }

  backupSheets.sort((a, b) => b.getName().localeCompare(a.getName()));
  Logger.log(`比較対象のバックアップシート: "${backupSheets[0].getName()}"`);
  return backupSheets[0];
}

/**
 * 【リファクタリング版】結果書き込み
 */
function writeResults(spreadsheet, differences) {
  let resultsSheet = spreadsheet.getSheetByName(CONFIG.RESULTS_SHEET_NAME);
  if (!resultsSheet) {
    resultsSheet = spreadsheet.insertSheet(CONFIG.RESULTS_SHEET_NAME);
  }

  resultsSheet.clear();

  // ヘッダーも定数を使って構築（整合性を保つため）
  const header = [[
    COLUMN_DEF.TEAM_NAME,
    '送付先メアド(TO)の変更後',
    '送付先メアド(CC)の変更後',
    '変更の種類',
    '備考'
  ]];

  resultsSheet.getRange("A1:E1").setValues(header).setFontWeight('bold');
  resultsSheet.setFrozenRows(1);

  if (differences.length > 0) {
    resultsSheet.getRange(2, 1, differences.length, 5).setValues(differences);
  }
  resultsSheet.autoResizeColumns(1, 5);
}

/**
 * 指定されたシートのバックアップを作成します。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheetToBackup - バックアップ対象のシート。
 * @returns {string} - 作成されたバックアップシート名。
 */
function backupSheet(sheetToBackup) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmm');
  const backupSheetName = `${sheetToBackup.getName()}${timestamp}`;
  sheetToBackup.copyTo(sheetToBackup.getParent()).setName(backupSheetName);
  Logger.log(`バックアップシートを作成しました: "${backupSheetName}"`);
  return backupSheetName;
}

/**
 * 指定されたIDとシート名からシートの全データを取得します。
 * @param {string} spreadsheetId - スプレッドシートのID。
 * @param {string} sheetName - シート名。
 * @returns {Array<Array<any>>} - シートのデータ（2次元配列）。
 */
function getSheetData(spreadsheetId, sheetName) {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`シート "${sheetName}" が見つかりません。`);
  }
  return sheet.getDataRange().getValues();
}

/**
 * 指定されたシートのデータをクリアし、新しいデータセットを書き込みます。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 更新対象のシート。
 * @param {Array<Array<any>>} data - 書き込む新しいデータ。
 */
function updateSheetWithData(sheet, data) {
  // ★★★ 変更点: ヘッダー（2行）以外の既存データをクリア ★★★
  if (sheet.getLastRow() > 2) { // 2行より大きい（データ行が存在する）場合のみクリア
    sheet.getRange(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn()).clearContent();
  }

  // ★★★ 変更点: 新しいデータを3行目から書き込み ★★★
  if (data.length > 0) {
    sheet.getRange(3, 1, data.length, data[0].length).setValues(data);
  }
}
