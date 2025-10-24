/**
 * utils/format.ts — Number formatting and validation utilities
 * ==============================================================
 * Purpose:
 *   - Type-safe number validation and conversion
 *   - Rounding and formatting helpers
 *   - Date formatting utilities
 * 
 * Dependencies: None
 * Called by: All calculation modules
 */

/**
 * Checks if a value is a finite number.
 * @param {any} v - Value to check
 * @returns {boolean} True if value is a finite number
 */
export function isNum(v: any): v is number {
  return typeof v === 'number' && isFinite(v);
}

/**
 * Converts value to number, returns null if invalid.
 * @param {any} v - Value to convert
 * @returns {number | null} Number or null
 */
export function toNum(v: any): number | null {
  if (v === null || v === undefined || v === '') {
    return null;
  }
  
  const n = Number(v);
  return isFinite(n) ? n : null;
}

/**
 * Converts value to integer, returns null if invalid.
 * @param {any} v - Value to convert
 * @returns {number | null} Integer or null
 */
export function toInt(v: any): number | null {
  const n = toNum(v);
  return n !== null ? Math.trunc(n) : null;
}

/**
 * Converts value to 4-digit fiscal year integer.
 * Accepts numbers or Date objects.
 * @param {any} v - Value to convert
 * @returns {number | null} Year or null
 */
export function toYear(v: any): number | null {
  if (v instanceof Date) {
    return v.getFullYear();
  }
  
  const n = toInt(v);
  if (n === null) return null;
  
  // Basic validation: should be a reasonable year
  if (n < 1900 || n > 2100) return null;
  
  return n;
}

/**
 * Rounds number to 1 decimal place.
 * @param {number | null} v - Value to round
 * @returns {number | null} Rounded value or null
 */
export function round1(v: any): number | null {
  const n = toNum(v);
  if (n === null) return null;
  return Math.round(n * 10) / 10;
}

/**
 * Rounds number to 2 decimal places.
 * @param {number | null} v - Value to round
 * @returns {number | null} Rounded value or null
 */
export function round2(v: any): number | null {
  const n = toNum(v);
  if (n === null) return null;
  return Math.round(n * 100) / 100;
}

/**
 * Rounds number to 3 decimal places.
 * @param {number | null} v - Value to round
 * @returns {number | null} Rounded value or null
 */
export function round3(v: any): number | null {
  const n = toNum(v);
  if (n === null) return null;
  return Math.round(n * 1000) / 1000;
}

/**
 * Rounds number to 4 decimal places (for CAGRs).
 * @param {number | null} v - Value to round
 * @returns {number | null} Rounded value or null
 */
export function round4(v: any): number | null {
  const n = toNum(v);
  if (n === null) return null;
  return Math.round(n * 10000) / 10000;
}

/**
 * Rounds number to N decimal places.
 * @param {number | null} v - Value to round
 * @param {number} decimals - Number of decimal places
 * @returns {number | null} Rounded value or null
 */
export function roundN(v: any, decimals: number): number | null {
  const n = toNum(v);
  if (n === null) return null;
  
  const factor = Math.pow(10, decimals);
  return Math.round(n * factor) / factor;
}

/**
 * Clamps a number within min/max bounds.
 * @param {number} v - Value to clamp
 * @param {number} min - Minimum value
 * @param {number} max - Maximum value
 * @returns {number} Clamped value
 */
export function clamp(v: number, min: number, max: number): number {
  return Math.min(Math.max(v, min), max);
}

/**
 * Applies the ≤0 → 0.01 floor rule for CAGR calculations.
 * Used to handle turnaround/negative scenarios.
 * @param {number | null} v - Value to floor
 * @returns {number | null} Floored value (0.01 minimum) or null
 */
export function floor01(v: any): number | null {
  const n = toNum(v);
  if (n === null) return null;
  return n <= 0 ? 0.01 : n;
}

/**
 * Safe division: returns null if denominator is zero or invalid.
 * @param {any} numerator - Numerator
 * @param {any} denominator - Denominator
 * @returns {number | null} Result or null
 */
export function safeDiv(numerator: any, denominator: any): number | null {
  const num = toNum(numerator);
  const den = toNum(denominator);
  
  if (num === null || den === null || den === 0) {
    return null;
  }
  
  return num / den;
}

/**
 * Formats a date as yyyy-mm-dd.
 * @param {Date} d - Date to format
 * @returns {string} Formatted date string
 */
export function formatDate(d: Date | null | undefined): string {
  if (!d || !(d instanceof Date)) return '';
  
  try {
    return Utilities.formatDate(
      d,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  } catch (e) {
    return '';
  }
}

/**
 * Formats a timestamp as yyyy-mm-dd hh:mm:ss.
 * @param {Date} d - Date to format
 * @returns {string} Formatted timestamp string
 */
export function formatTimestamp(d: Date | null | undefined): string {
  if (!d || !(d instanceof Date)) return '';
  
  try {
    return Utilities.formatDate(
      d,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );
  } catch (e) {
    return '';
  }
}

/**
 * Returns first non-null value from array, or null if all null.
 * @param {any[]} values - Array of values to check
 * @returns {any | null} First non-null value or null
 */
export function firstNonNull(values: any[]): any | null {
  for (const v of values) {
    if (v !== null && v !== undefined && v !== '') {
      return v;
    }
  }
  return null;
}

/**
 * Safely accesses a property from an object, returns null if missing.
 * @param {any} obj - Object to access
 * @param {string | number} key - Key to access
 * @returns {any | null} Value or null
 */
export function maybe(obj: any, key: string | number): any | null {
  if (obj == null) return null;
  
  if (Array.isArray(obj)) {
    const idx = Number(key);
    return (idx >= 0 && idx < obj.length) ? obj[idx] : null;
  }
  
  if (typeof obj === 'object') {
    return (key in obj) ? obj[key] : null;
  }
  
  return null;
}

/**
 * Converts empty strings to null for cleaner spreadsheet output.
 * @param {any} v - Value to check
 * @returns {any} Value or null if empty string
 */
export function emptyToNull(v: any): any {
  return v === '' ? null : v;
}

/**
 * Converts null/undefined to empty string for spreadsheet output.
 * @param {any} v - Value to check
 * @returns {any} Value or empty string
 */
export function nullToEmpty(v: any): any {
  return (v === null || v === undefined) ? '' : v;
}

/**
 * Returns blank ('') for invalid numeric cells.
 * Use this for per-share metrics that should show blank instead of 0.
 * @param {number | null} v - Value to format
 * @returns {number | string} Value or blank string
 */
export function blankIfNull(v: number | null): number | string {
  return v !== null ? v : '';
}

