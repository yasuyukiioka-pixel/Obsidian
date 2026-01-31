/**
 * @file Code.js
 * @brief Main script file containing the primary workflow orchestration and UI menu setup.
 * @author yasuyukiioka-pixel
 */

/**
 * @function transferData
 * @brief Main function to orchestrate the multi-stage data transfer and transformation process.
 *
 * This function executes a 4-stage data processing pipeline:
 * 1. Reads raw data from the source sheet.
 * 2. Applies initial column mapping and transformation.
 * 3. Removes empty rows while preserving multi-line CC data.
 * 4. Consolidates multi-line entries for each team into a single row.
 * 5. Removes teams that have no TO or CC email addresses.
 * The output of each stage is written to a separate sheet for debugging and transparency.
 *
 * @throws {Error} If any stage of the process fails, the error is logged.
 */
function transferData() {
  try {
    Logger.log('Starting data transfer process...');

    // Read headers from source spreadsheet (if needed)
    let sourceHeaders = [];
    if (CONFIG.INCLUDE_HEADERS) {
      sourceHeaders = readSourceHeaders();
      Logger.log(`Read ${sourceHeaders.length} header rows from source spreadsheet`);
    }

    // Read data from source spreadsheet
    const sourceData = readSourceData();
    Logger.log(`Read ${sourceData.length} data rows from source spreadsheet`);

    // Transform headers to target format (if needed)
    let transformedHeaders = [];
    if (CONFIG.INCLUDE_HEADERS && sourceHeaders.length > 0) {
      transformedHeaders = transformHeaders(sourceHeaders);
      Logger.log(`Transformed headers to ${transformedHeaders.length} rows`);
    }

    // Transform data to target format using 4-stage process
    Logger.log('Stage 1: Initial transformation with column mapping');
    const stage1Data = transformData(sourceData);
    Logger.log(`Stage 1 completed: ${stage1Data.length} rows`);

    // Write Stage 1 results to "1.出力" sheet
    writeToTargetSheet(transformedHeaders, stage1Data, '1.出力');
    Logger.log('Stage 1 data written to "1.出力" sheet');

    Logger.log('Stage 2: Remove empty rows while preserving CC data');
    const stage2Data = removeEmptyRowsStage2(stage1Data);
    Logger.log(`Stage 2 completed: ${stage2Data.length} rows`);

    // Write Stage 2 results to "2.空白なし" sheet
    writeToTargetSheet(transformedHeaders, stage2Data, '2.空白なし');
    Logger.log('Stage 2 data written to "2.空白なし" sheet');

    Logger.log('Stage 3: Consolidate multiple rows');
    const stage3Data = consolidateMultipleRows(stage2Data);
    Logger.log(`Stage 3 completed: ${stage3Data.length} rows`);

    // Write Stage 3 results to "3.空白・複数行なし" sheet
    writeToTargetSheet(transformedHeaders, stage3Data, '3.空白・複数行なし');
    Logger.log('Stage 3 data written to "3.空白・複数行なし" sheet');

    Logger.log('Stage 4: Remove rows without To/CC email data');
    const stage4Data = removeRowsWithoutToOrCCStage4(stage3Data);
    Logger.log(`Stage 4 completed: ${stage4Data.length} rows`);

    // Write Stage 4 results to "4.To・CCなし行削除" sheet
    writeToTargetSheet(transformedHeaders, stage4Data, '4.To・CCなし行削除');
    Logger.log('Stage 4 data written to "4.To・CCなし行削除" sheet');

    Logger.log('All 4-stage data transfer completed successfully');

    Logger.log('データ転送が完了しました！');
  } catch (error) {
    Logger.log('Error during data transfer: ' + error.toString());
  }
}

/**
 * @function allexec
 * @brief Executes the complete data pipeline: transfer, update, and compare.
 * A convenience function for running the entire workflow in sequence.
 */
function allexec() {
  transferData()
  updateMailSettings_refactored()
  compareAndUpdateResults_refactored()
}

/**
 * @function scheduledDuplicateCheck
 * @brief Wrapper function for the time-driven trigger.
 * Runs the duplicate check non-interactively.
 */
function scheduledDuplicateCheck() {
  checkForDuplicatesAndNotify(false);
}

/**
 * @function manualDuplicateCheck
 * @brief Wrapper function for the menu item.
 * Runs the duplicate check interactively (showing alerts).
 */
function manualDuplicateCheck() {
  checkForDuplicatesAndNotify(true);
}

/**
 * @function onOpen
 * @brief Creates a custom menu in the Google Sheets UI when the spreadsheet is opened.
 * This menu provides easy access to the main functions of the script.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('データ転送ツール')
    .addItem('データを転送', 'transferData')
    .addItem('リストを作成', 'updateMailSettings_refactored')
    .addItem('リストを比較', 'compareAndUpdateResults_refactored')
    .addSeparator()
//    .addItem('新規登録の重複チェック', 'manualDuplicateCheck')
    .addItem('重複チェック (前回以降)', 'runIncrementalDuplicateCheck')
    .addSeparator()
    .addItem('一括実行', 'allexec')
    .addToUi();
}