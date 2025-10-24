/***************************************************************
 * Utils.gs — shared helpers used across Screener project
 * -------------------------------------------------------------
 * Keep this file lightweight and global — one source of truth
 * for helpers like isNum(), toNum(), round1(), etc.
 ***************************************************************/

/**
 * Check if a value is a finite number
 * @param {*} v
 * @return {boolean}
 */
function isNum(v) {
  return typeof v === 'number' && isFinite(v);
}

/**
 * Convert to number if valid, else return null
 * @param {*} v
 * @return {number|null}
 */
function toNum(v) {
  return (typeof v === 'number' && isFinite(v)) ? v : null;
}

/**
 * Round to 1 decimal place
 * @param {number} v
 * @return {number|null}
 */
function round1(v) {
  return (typeof v === 'number' && isFinite(v)) ? Math.round(v * 10) / 10 : null;
}

/**
 * Round to N decimals
 * @param {number} v
 * @param {number} n
 * @return {number|null}
 */
function roundN(v, n) {
  if (!isNum(v)) return null;
  const p = Math.pow(10, n);
  return Math.round(v * p) / p;
}

/**
 * Safe access: return value if key exists, else null
 * @param {Array|Object} obj
 * @param {string|number} key
 * @return {*|null}
 */
function maybe(obj, key) {
  if (obj == null) return null;
  if (Array.isArray(obj)) return (key < obj.length) ? obj[key] : null;
  if (typeof obj === 'object') return (key in obj) ? obj[key] : null;
  return null;
}

/**
 * Convert value to a 4-digit fiscal year integer
 * Accepts numbers or Date objects
 * @param {*} v
 * @return {number|null}
 */
function toYear(v) {
  if (v instanceof Date) return v.getFullYear();
  const n = +v;
  return isFinite(n) ? Math.round(n) : null;
}

/**
 * Clamp a number within min/max bounds
 * @param {number} v
 * @param {number} min
 * @param {number} max
 */
function clamp(v, min, max) {
  return Math.min(Math.max(v, min), max);
}

/**
 * Format a date as yyyy-mm-dd
 * @param {Date} d
 * @return {string}
 */
function fmtDate(d) {
  if (!(d instanceof Date)) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Simple timestamped log message (appears in execution logs)
 * @param {string} msg
 */
function log_(msg) {
  console.log(`[${new Date().toISOString()}] ${msg}`);
}
