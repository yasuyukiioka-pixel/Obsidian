/**
 * @file utils.js
 * @brief This file contains utility functions.
 */

/**
 * Convert column letter to index (A=0, B=1, etc.)
 * @param {string} columnLetter - Column letter (e.g., 'A', 'B', 'AA')
 * @return {number} Column index, or -1 if invalid
 */
function columnLetterToIndex(columnLetter) {
  if (!columnLetter || typeof columnLetter !== 'string') {
    return -1;
  }

  let result = 0;
  for (let i = 0; i < columnLetter.length; i++) {
    const char = columnLetter.charCodeAt(i) - 65; // A=0, B=1, etc.
    if (char < 0 || char > 25) {
      return -1; // Invalid character
    }
    result = result * 26 + char + 1;
  }
  return result - 1;
}

/**
 * Convert column index to letter (0=A, 1=B, etc.)
 * @param {number} columnIndex - Column index
 * @return {string} Column letter
 */
function columnIndexToLetter(columnIndex) {
  let result = '';
  while (columnIndex >= 0) {
    result = String.fromCharCode(65 + (columnIndex % 26)) + result;
    columnIndex = Math.floor(columnIndex / 26) - 1;
  }
  return result;
}


/**
 * Helper function to check if a value looks like a date
 * @param {*} value - Value to check
 * @return {boolean} True if the value appears to be a date
 */
function isDateLike(value) {
  if (!value) return false;

  if (value instanceof Date) return true;

  if (typeof value === 'string') {
    const datePatterns = [
      /^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/,
      /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/,
      /^\d{4}\d{2}\d{2}$/
    ];

    return datePatterns.some(pattern => pattern.test(value.toString().trim()));
  }

  try {
    const date = new Date(value);
    return !isNaN(date.getTime());
  } catch (error) {
    return false;
  }
}

/**
 * Format date values
 * @param {*} value - Date value to format
 * @return {string} Formatted date string
 */
function formatDate(value) {
  if (!value) return '';

  try {
    const date = new Date(value);
    if (isNaN(date.getTime())) return value.toString();

    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  } catch (error) {
    return value.toString();
  }
}

/**
 * Format number values
 * @param {*} value - Number value to format
 * @return {number|string} Formatted number
 */
function formatNumber(value) {
  if (!value) return '';

  const num = parseFloat(value);
  if (isNaN(num)) return value.toString();

  return Math.round(num * 100) / 100;
}
