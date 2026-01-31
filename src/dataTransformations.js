/**
 * @file dataTransformations.js
 * @brief This file contains functions for transforming and cleaning data.
 */

/**
 * Stage 2: Remove empty rows while preserving CC data
 * 2.空白なし - Remove rows that are completely empty but keep rows with CC data
 * @param {Array} data - 2D array of transformed data
 * @return {Array} 2D array with empty rows removed
 */
function removeEmptyRowsStage2(data) {
  if (!data || data.length === 0) {
    return [];
  }

  const filteredData = data.filter(row => {
    // Check if row has any meaningful data
    return row.some(cell => {
      if (!cell) return false;
      const cellStr = cell.toString();
      // Preserve multi-line content including CC email lists
      return cellStr.replace(/\s/g, '') !== '';
    });
  });

  Logger.log(`Stage 2: Filtered ${data.length} rows to ${filteredData.length} rows (removed empty rows)`);
  return filteredData;
}

/**
 * Stage 3: Consolidate multiple rows with the same team name into a single row.
 * Merges email addresses (To/CC) with newline separators.
 * @param {Array} data - 2D array of cleaned data from Stage 2.
 * @return {Array} 2D array with consolidated rows.
 */
function consolidateMultipleRows(data) {
  if (!data || data.length === 0) {
    return [];
  }

  const consolidated = [];
  let currentTeamRows = [];

  function processTeam() {
    if (currentTeamRows.length > 0) {
      const consolidatedRow = consolidateTeamRows(currentTeamRows);
      consolidated.push(consolidatedRow);
    }
  }

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const teamName = (row[0] || '').toString().trim();

    if (teamName !== '') {
      // New team starts, process the previous one
      processTeam();
      // Start a new team group
      currentTeamRows = [row];
    } else {
      // This is a continuation row for the current team
      if (currentTeamRows.length > 0) {
        currentTeamRows.push(row);
      }
    }
  }

  // Process the last team group
  processTeam();

  Logger.log(`Consolidated ${data.length} rows into ${consolidated.length} rows by team name.`);
  return consolidated;
}

/**
 * Stage 4: Remove rows where both To and CC columns are empty
 * 4.To・CCなし行削除 - Filter out teams with no email recipients
 * @param {Array} data - 2D array of data from stage 3
 * @return {Array} 2D array with rows filtered by To/CC presence
 */
function removeRowsWithoutToOrCCStage4(data) {
  if (!data || data.length === 0) {
    return [];
  }

  const mappingRules = readColumnMappingRules();
  const toRule = mappingRules.find(r => r.description.includes('To'));
  const ccRule = mappingRules.find(r => r.description.includes('CC'));

  if (!toRule || !ccRule) {
      Logger.log('Warning: To/CC mapping rules not found. Cannot perform Stage 4 filtering.');
      return data;
  }

  const toIndex = toRule.targetIndex;
  const ccIndex = ccRule.targetIndex;

  const filteredData = data.filter(row => {
    const teamName = (row[0] || '').toString().trim();
    if (teamName === '') {
      return false;
    }

    const toEmail = (row[toIndex] || '').toString().trim();
    const hasToEmail = toEmail !== '';

    const ccEmail = (row[ccIndex] || '').toString().trim();
    const hasCCEmail = ccEmail !== '';

    const shouldKeep = hasToEmail || hasCCEmail;

    if (!shouldKeep) {
      Logger.log(`Filtering out team "${teamName}" - no To/CC email data`);
    }

    return shouldKeep;
  });

  Logger.log(`Stage 4: Filtered ${data.length} rows to ${filteredData.length} rows (removed rows without To/CC emails)`);
  return filteredData;
}

/**
 * Consolidate multiple rows for the same team into a single row
 * @param {Array} teamRows - Array of rows belonging to the same team
 * @return {Array} Single consolidated row
 */
function consolidateTeamRows(teamRows) {
  if (!teamRows || teamRows.length === 0) return [];
  if (teamRows.length === 1) return [...teamRows[0]];

  const consolidatedRow = [...teamRows[0]];

  for (let i = 1; i < teamRows.length; i++) {
    const currentRow = teamRows[i];

    for (let colIndex = 0; colIndex < currentRow.length; colIndex++) {
      const cellValue = (currentRow[colIndex] || '').toString().trim();

      if (cellValue !== '') {
        const existingValue = (consolidatedRow[colIndex] || '').toString().trim();

        if (cellValue.includes('@')) {
          if (existingValue === '') {
            consolidatedRow[colIndex] = cellValue;
          } else {
            const existingEmails = existingValue.split('\n').map(e => e.trim());
            if (!existingEmails.includes(cellValue)) {
              consolidatedRow[colIndex] = existingValue + '\n' + cellValue;
            }
          }
        } else if (existingValue === '') {
          consolidatedRow[colIndex] = cellValue;
        }
      }
    }
  }

  return consolidatedRow;
}

/**
 * Transform data from source format to target format using dynamic column mapping
 * Column mapping rules are read from the mapping spreadsheet
 * @param {Array} sourceData - 2D array of source data
 * @return {Array} 2D array of transformed data
 */
function transformData(sourceData) {
  if (!sourceData || sourceData.length === 0) {
    return [];
  }

  const mappingRules = readColumnMappingRules();

  if (mappingRules.length === 0) {
    Logger.log('Warning: No column mapping rules found. Data will not be transformed.');
    return [];
  }

  const maxTargetIndex = Math.max(...mappingRules.map(rule => rule.targetIndex));
  const targetColumnCount = maxTargetIndex + 1;

  const transformedData = [];

  for (let i = 0; i < sourceData.length; i++) {
    const sourceRow = sourceData[i];

    if (sourceRow.every(cell => cell === '')) {
      continue;
    }

    const transformedRow = new Array(targetColumnCount).fill('');

    mappingRules.forEach(rule => {
      if (sourceRow.length > rule.sourceIndex && sourceRow[rule.sourceIndex] !== undefined) {
        let value = sourceRow[rule.sourceIndex] !== null ? sourceRow[rule.sourceIndex] : '';

        if (value && isDateLike(value)) {
          value = formatDate(value);
        } else if (value && typeof value === 'number') {
          value = formatNumber(value);
        }

        transformedRow[rule.targetIndex] = value;
      }
    });

    const hasData = transformedRow.some(cell => {
      if (!cell) return false;
      const cellStr = cell.toString();
      return cellStr.replace(/\s/g, '') !== '';
    });

    if (hasData) {
      transformedData.push(transformedRow);
    }
  }

  Logger.log(`Transformed ${transformedData.length} rows using ${mappingRules.length} mapping rules`);
  return transformedData;
}
