/**
 * Config.js
 * Change your settings here.
 */

// Default salary cycle start day (fallback if not set in sheet)
const DEFAULT_CYCLE_START_DAY = 9;

// Read cycle start day from _Config sheet, fallback to default
function getCycleStartDay() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('_Config');
        if (sheet) {
            const val = sheet.getRange('B1').getValue();
            if (val && parseInt(val, 10) >= 1 && parseInt(val, 10) <= 28) {
                return parseInt(val, 10);
            }
        }
    } catch (e) { /* fallback */ }
    return DEFAULT_CYCLE_START_DAY;
}

// For backward compatibility — used as a getter
var CYCLE_START_DAY = DEFAULT_CYCLE_START_DAY;

// Initialize on script load
function initConfig() {
    CYCLE_START_DAY = getCycleStartDay();
}
