/**
 * utils/spinner.ts — UI feedback and timing wrapper
 * ===================================================
 * Purpose:
 *   - Wrap long operations with user feedback
 *   - Show start confirmation and elapsed time
 *   - Centralized error handling with UI alerts
 * 
 * Dependencies: None
 * Called by: Menu entrypoints in code.ts
 */

/**
 * Wraps a function with UI spinner: shows start alert, runs function,
 * then displays completion time.
 * @param {Function} fn - Function to execute
 * @param {string} label - Label to display to user
 * @returns {any} Result of the function
 * @throws {Error} Re-throws any error after displaying to user
 */
export function withUiSpinner<T>(fn: () => T, label: string): T {
  const ui = SpreadsheetApp.getUi();
  
  // Show start confirmation
  ui.alert(`${label}\n\nClick OK to start.`);
  
  const startTime = Date.now();
  
  try {
    // Execute the function
    const result = fn();
    
    // Calculate elapsed time
    const elapsedMs = Date.now() - startTime;
    const elapsedSec = (elapsedMs / 1000).toFixed(1);
    
    // Log success
    Logger.log(`✓ ${label} — completed in ${elapsedSec}s`);
    
    // Show success alert
    ui.alert(`Done.\nElapsed: ${elapsedSec}s`);
    
    return result;
    
  } catch (e: any) {
    // Calculate elapsed time even on error
    const elapsedMs = Date.now() - startTime;
    const elapsedSec = (elapsedMs / 1000).toFixed(1);
    
    // Log error with stack trace
    const errorMsg = e && e.stack ? e.stack : String(e);
    Logger.log(`✗ ${label} — failed after ${elapsedSec}s\n${errorMsg}`);
    
    // Show error alert
    const displayMsg = e && e.message ? e.message : String(e);
    ui.alert(`Error:\n${displayMsg}`);
    
    // Re-throw for caller to handle if needed
    throw e;
  }
}

/**
 * Wraps a function with timing only (no UI alerts).
 * Logs start, end, and elapsed time.
 * @param {Function} fn - Function to execute
 * @param {string} label - Label for logs
 * @returns {any} Result of the function
 */
export function withTiming<T>(fn: () => T, label: string): T {
  Logger.log(`[START] ${label}`);
  const startTime = Date.now();
  
  try {
    const result = fn();
    const elapsedMs = Date.now() - startTime;
    Logger.log(`[DONE] ${label} — ${elapsedMs}ms`);
    return result;
    
  } catch (e: any) {
    const elapsedMs = Date.now() - startTime;
    const errorMsg = e && e.stack ? e.stack : String(e);
    Logger.log(`[ERROR] ${label} — ${elapsedMs}ms\n${errorMsg}`);
    throw e;
  }
}

/**
 * Shows a simple info alert to the user.
 * @param {string} message - Message to display
 */
export function showInfo(message: string): void {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log('showInfo failed: ' + e);
  }
}

/**
 * Shows an error alert to the user.
 * @param {string} title - Error title
 * @param {string} message - Error message
 */
export function showError(title: string, message: string): void {
  try {
    SpreadsheetApp.getUi().alert(`${title}\n\n${message}`);
  } catch (e) {
    Logger.log('showError failed: ' + e);
  }
}

/**
 * Prompts user for input with a dialog.
 * @param {string} title - Dialog title
 * @param {string} prompt - Prompt message
 * @returns {string | null} User input or null if cancelled
 */
export function promptUser(title: string, prompt: string): string | null {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const response = ui.prompt(title, prompt, ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return null;
    }
    
    const text = response.getResponseText();
    return text ? text.trim() : null;
    
  } catch (e) {
    Logger.log('promptUser failed: ' + e);
    return null;
  }
}

/**
 * Shows a confirmation dialog.
 * @param {string} title - Dialog title
 * @param {string} message - Confirmation message
 * @returns {boolean} True if user clicked Yes
 */
export function confirmUser(title: string, message: string): boolean {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const response = ui.alert(title, message, ui.ButtonSet.YES_NO);
    return response === ui.Button.YES;
    
  } catch (e) {
    Logger.log('confirmUser failed: ' + e);
    return false;
  }
}

/**
 * Logs a timestamped message.
 * @param {string} message - Message to log
 */
export function logMessage(message: string): void {
  const timestamp = new Date().toISOString();
  console.log(`[${timestamp}] ${message}`);
}

/**
 * Checks if a required function exists in the global scope.
 * Throws helpful error if missing.
 * @param {string} functionName - Name of function to check
 * @throws {Error} If function not found
 */
export function requireFunction(functionName: string): void {
  const globalScope = this as any;
  const fn = globalScope[functionName];
  
  if (typeof fn !== 'function') {
    throw new Error(
      `Required function "${functionName}" not found. ` +
      'Make sure the corresponding file is included in the project.'
    );
  }
}

