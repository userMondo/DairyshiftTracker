/**
 * Code.js
 * Google Apps Script Backend for Dairy Shift Tracker
 */

// Format date to UTC+07:00
function formatTime(date) {
    return Utilities.formatDate(new Date(date), "GMT+07:00", "yyyy-MM-dd HH:mm:ss");
}

function getSheetNameOffset(date, offsetMonths) {
    const timezone = "GMT+07:00";
    let baseDate = date || new Date();
    const day = parseInt(Utilities.formatDate(baseDate, timezone, "dd"), 10);
    let cycleMonth = parseInt(Utilities.formatDate(baseDate, timezone, "MM"), 10) - 1; // 0-indexed
    let cycleYear = parseInt(Utilities.formatDate(baseDate, timezone, "yyyy"), 10);

    // Salary cycle: if day < CYCLE_START_DAY, we're still in previous month's cycle
    if (day < CYCLE_START_DAY) {
        cycleMonth -= 1;
    }

    // Apply offset
    cycleMonth += offsetMonths;

    // Normalize month/year
    while (cycleMonth < 0) { cycleMonth += 12; cycleYear -= 1; }
    while (cycleMonth > 11) { cycleMonth -= 12; cycleYear += 1; }

    return `${cycleMonth + 1}/${String(cycleYear).slice(-2)}`;
}

// Returns the cycle sheet name for a given YYYY-MM-DD date string
function getSheetNameForDate(dateStr) {
    const parts = dateStr.split('-');
    const day = parseInt(parts[2], 10);
    let month = parseInt(parts[1], 10); // 1-indexed
    let year = parseInt(parts[0], 10);

    if (day < CYCLE_START_DAY) {
        month -= 1;
        if (month < 1) { month = 12; year -= 1; }
    }

    return `${month}/${String(year).slice(-2)}`;
}

function getOrCreateSheet(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        sheet = ss.insertSheet(sheetName, 0); // Insert at beginning
        sheet.appendRow(["ID", "Date", "Start Time", "End Time", "Total Hours", "Note"]);
        sheet.setFrozenRows(1);
    }
    return sheet;
}

function doGet(e) {
    // Check if API request
    if (e.parameter.action) {
        return handleApiRequest(e.parameter.action, e);
    }

    // Default: Serve Web App UI
    return HtmlService.createHtmlOutputFromFile('DairyShift')
        .setTitle('DairyShift')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
    // Handle POST requests (Actions)
    const action = e.parameter.action;
    return handleApiRequest(action, e);
}

function handleApiRequest(action, e) {
    let result = {};

    try {
        // Refresh config from sheet on each request
        initConfig();

        if (action === 'getStatus') {
            result = getShiftStatus();
        } else if (action === 'getConfig') {
            result = getAppConfig();
        } else if (action === 'updateConfig') {
            const data = JSON.parse(e.postData.contents);
            result = updateAppConfig(data);
        } else if (action === 'getHistory') {
            const offset = (e.parameter && e.parameter.offset) ? parseInt(e.parameter.offset, 10) : 0;
            result = getShiftHistory(offset);
        } else if (action === 'start') {
            const data = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
            result = toggleShift('start', data.timestamp);
        } else if (action === 'end') {
            const data = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
            result = toggleShift('end', data.timestamp);
        } else if (action === 'delete') {
            if (!e.parameter.id) throw new Error("Missing ID for deletion");
            result = deleteShift(e.parameter.id);
        } else if (action === 'edit') {
            const data = JSON.parse(e.postData.contents);
            result = editShift(data);
        } else if (action === 'addManual') {
            const data = JSON.parse(e.postData.contents);
            result = addManualShift(data);
        } else if (action === 'archive') {
            result = archivePreviousMonth();
        } else {
            result = { error: "Invalid action" };
        }
    } catch (err) {
        result = { error: err.message };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Checks for active shift.
 * Looks in current month sheet first, then previous month (for overnight shifts).
 */
function getShiftStatus() {
    const now = new Date();
    const currentSheetName = getSheetNameOffset(now, 0);

    // Check current month
    let status = checkSheetForActiveShift(currentSheetName);
    if (status.active) return status;

    // Check previous month (handle month boundaries)
    const prevSheetName = getSheetNameOffset(now, -1);

    if (prevSheetName !== currentSheetName) {
        status = checkSheetForActiveShift(prevSheetName);
        if (status.active) return status;
    }

    return { active: false };
}

function checkSheetForActiveShift(sheetName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return { active: false };

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { active: false };

    // Scan backwards from the last row to find any active shift
    // Limit to last 100 rows for performance
    const numRows = Math.min(100, lastRow - 1);
    const startRow = lastRow - numRows + 1;

    // Columns: 1:ID, 2:Date, 3:Start Time, 4:End Time
    const data = sheet.getRange(startRow, 1, numRows, 4).getDisplayValues();

    for (let i = data.length - 1; i >= 0; i--) {
        const row = data[i];
        const date = row[1];
        const startTime = row[2];
        const endTime = row[3];

        if (date && startTime && !endTime) {
            return { active: true, startTime: date + " " + startTime };
        }
    }

    return { active: false };
}

/**
 * Toggles the shift state.
 * @param {string} action - "start" or "end"
 */
function toggleShift(action, clientTimestamp) {
    const timezone = "GMT+07:00";
    const now = clientTimestamp ? new Date(clientTimestamp) : new Date();
    const dateStr = Utilities.formatDate(now, timezone, "yyyy-MM-dd");
    const timeStr = Utilities.formatDate(now, timezone, "HH:mm:ss");
    const fullTimestamp = dateStr + " " + timeStr;
    const sheetName = getSheetNameOffset(now, 0);

    if (action === 'start') {
        const sheet = getOrCreateSheet(sheetName);

        // Check local active state first (rudimentary check on last row of THIS sheet)
        // ideally we trust getShiftStatus, but here we just append.
        // For robustness, users should not be able to start if status says active.

        // Double check we don't have an active shift globally
        const status = getShiftStatus();
        if (status.active) throw new Error("Shift already active!");

        // ID Generation: Simple timestamp + random
        const id = Utilities.formatDate(now, timezone, "yyyyMMddHHmmss");

        // Append: [ID, Date, Start Time, End Time (Empty), Total Hours (Empty), Note (Empty)]
        sheet.appendRow([id, dateStr, timeStr, "", "", ""]);

        return { success: true, state: 'active', startTime: fullTimestamp };

    } else if (action === 'end') {
        // Find where the active shift is (could be previous month)
        let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        let activeRowIndex = -1;

        // 1. Try current month
        // Check End Time column (Column 4) backwards
        if (sheet && sheet.getLastRow() > 1) {
            const numRows = Math.min(100, sheet.getLastRow() - 1);
            const startRow = sheet.getLastRow() - numRows + 1;
            const data = sheet.getRange(startRow, 4, numRows).getValues().flat();
            for (let i = data.length - 1; i >= 0; i--) {
                if (data[i] === "") {
                    activeRowIndex = startRow + i;
                    break;
                }
            }
        }

        // 2. If not found, try previous month
        if (activeRowIndex === -1) {
            const prevSheetName = getSheetNameOffset(now, -1);
            const prevSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(prevSheetName);

            if (prevSheet && prevSheet.getLastRow() > 1) {
                const numRows = Math.min(100, prevSheet.getLastRow() - 1);
                const startRow = prevSheet.getLastRow() - numRows + 1;
                const data = prevSheet.getRange(startRow, 4, numRows).getValues().flat(); // End Time col
                for (let i = data.length - 1; i >= 0; i--) {
                    if (data[i] === "") {
                        sheet = prevSheet;
                        activeRowIndex = startRow + i;
                        break;
                    }
                }
            }
        }

        if (activeRowIndex === -1) throw new Error("No active shift to end.");

        // Update End Time (Col 4)
        sheet.getRange(activeRowIndex, 4).setValue(timeStr);

        // Set Total Hours formula (Col 5)
        const formula = `=TEXT(MOD(D${activeRowIndex}-C${activeRowIndex},1)*24,"0.00")`;
        sheet.getRange(activeRowIndex, 5).setFormula(formula);

        // Read back calculated value for API response
        SpreadsheetApp.flush();
        const hours = sheet.getRange(activeRowIndex, 5).getDisplayValue();

        return { success: true, state: 'inactive', totalHours: hours };
    }
}

/**
 * Deletes a shift by ID.
 * Searches current and previous month.
 */
function deleteShift(id) {
    const now = new Date();
    const currentSheetName = getSheetNameOffset(now, 0);
    const prevSheetName = getSheetNameOffset(now, -1);

    // Try current
    if (deleteFromSheet(currentSheetName, id)) return { success: true };
    // Try previous
    if (prevSheetName !== currentSheetName) {
        if (deleteFromSheet(prevSheetName, id)) return { success: true };
    }

    throw new Error("Shift ID not found.");
}

/**
 * Edits a shift.
 * @param {Object} data - { id, date, startTime, endTime, note }
 */
function editShift(data) {
    const { id, date, startTime, endTime, note } = data;
    if (!id) throw new Error("Missing ID for edit");

    const now = new Date();
    const currentSheetName = getSheetNameOffset(now, 0);
    const prevSheetName = getSheetNameOffset(now, -1);

    // Try current
    if (updateSheetRow(currentSheetName, id, date, startTime, endTime, note)) return { success: true };
    // Try previous
    if (prevSheetName !== currentSheetName) {
        if (updateSheetRow(prevSheetName, id, date, startTime, endTime, note)) return { success: true };
    }

    throw new Error("Shift ID not found for edit.");
}

function updateSheetRow(sheetName, id, date, startTime, endTime, note) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return false;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return false;

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const rowIndex = ids.findIndex(rowId => String(rowId) === String(id));

    if (rowIndex !== -1) {
        const row = rowIndex + 2;

        // Columns: 2:Date, 3:Start Time, 4:End Time
        sheet.getRange(row, 2).setValue(date);
        sheet.getRange(row, 3).setValue(startTime);
        sheet.getRange(row, 4).setValue(endTime || ""); // Allow clearing

        if (sheet.getMaxColumns() < 6) {
            sheet.insertColumnAfter(5);
            sheet.getRange(1, 6).setValue("Note");
        }
        sheet.getRange(row, 6).setValue(note || "");

        // Recalculate Duration if end time exists
        if (endTime) {
            const formula = `=TEXT(MOD(D${row}-C${row},1)*24,"0.00")`;
            sheet.getRange(row, 5).setFormula(formula);
        } else {
            sheet.getRange(row, 5).setValue("");
        }

        return true;
    }
    return false;
}

function deleteFromSheet(sheetName, id) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return false;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return false;

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const rowIndex = ids.findIndex(rowId => String(rowId) === String(id));

    if (rowIndex !== -1) {
        sheet.deleteRow(rowIndex + 2);
        return true;
    }
    return false;
}

/**
 * Gets the shifts for history from a specific month offset
 * - 0 = current month, -1 = last month, etc.
 */
function getShiftHistory(offset = 0) {
    const now = new Date();
    // Calculate the target month based on the offset safely
    const sheetName = getSheetNameOffset(now, offset);

    // Instead of getOrCreateSheet, just get the sheet. If it doesn't exist, return empty array.
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    let history = [];

    if (sheet) {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
            // Get all rows in the month (or a larger limit if needed, e.g. 50)
            // Let's get up to 50 rows for history view, recent first.
            const numRows = Math.min(50, lastRow - 1);
            const startRow = lastRow - numRows + 1;

            // Get values and reverse to show newest first
            // Columns: 1:ID, 2:Date, 3:StartTime, 4:EndTime, 5:TotalHours, 6:Note
            let data = [];
            if (sheet.getMaxColumns() >= 6) {
                data = sheet.getRange(startRow, 1, numRows, 6).getValues().reverse();
            } else {
                const temp = sheet.getRange(startRow, 1, numRows, 5).getValues().reverse();
                data = temp.map(r => [...r, ""]);
            }

            history = data.map(row => {
                // Format date: handle both Date objects and strings
                let date = row[1];
                if (date instanceof Date) {
                    date = Utilities.formatDate(date, "GMT+07:00", "yyyy-MM-dd");
                } else {
                    date = String(date);
                }

                // Format times: handle both Date objects and strings
                let startTime = row[2];
                if (startTime instanceof Date) {
                    startTime = Utilities.formatDate(startTime, "GMT+07:00", "HH:mm:ss");
                } else {
                    startTime = String(startTime);
                }

                let endTime = row[3];
                if (endTime instanceof Date) {
                    endTime = Utilities.formatDate(endTime, "GMT+07:00", "HH:mm:ss");
                } else {
                    endTime = String(endTime);
                }

                // Reconstruct End Date for display if overnight
                let endDateStr = date;
                if (endTime && startTime) {
                    const startDt = new Date(date + " " + startTime); // Validation only
                    const endDt = new Date(date + " " + endTime);
                    if (endDt < startDt) {
                        // It's next day
                        // We can just append "(+1)" or something, but frontend expects YYYY-MM-DD HH:mm:ss
                        // Let's actually calculate the next day date string
                        const nextDay = new Date(endDt);
                        nextDay.setDate(nextDay.getDate() + 1);
                        endDateStr = Utilities.formatDate(nextDay, "GMT+07:00", "yyyy-MM-dd");
                    }
                }

                return {
                    id: row[0],
                    start: date + " " + startTime,
                    end: endTime ? (endDateStr + " " + endTime) : "",
                    hours: String(row[4]),
                    note: row[5] || ""
                };
            });

            // Sort history descending chronologically (newest first)
            history.sort((a, b) => new Date(b.start).getTime() - new Date(a.start).getTime());
        }
    }

    // Build cycle range label from sheet name (e.g. "3/26" -> "3/8 - 4/7")
    const sheetParts = sheetName.split('/');
    const cycleM = parseInt(sheetParts[0], 10);
    let endM = cycleM + 1;
    if (endM > 12) endM = 1;
    const endDay = CYCLE_START_DAY - 1;
    const cycleLabel = `${cycleM}/${CYCLE_START_DAY} - ${endM}/${endDay}`;

    return {
        monthLabel: cycleLabel,
        data: history
    };
}

/**
 * Adds a manual shift record.
 * @param {Object} data - { date, startTime, endTime }
 */
function addManualShift(data) {
    const { date, startTime, endTime, note } = data; // date is YYYY-MM-DD
    if (!date || !startTime || !endTime) throw new Error("Missing required fields");

    const sheetName = getSheetNameForDate(date);
    const sheet = getOrCreateSheet(sheetName);

    // Generate ID
    const timezone = "GMT+07:00";
    const now = new Date();
    const id = Utilities.formatDate(now, timezone, "yyyyMMddHHmmss") + Math.floor(Math.random() * 1000);

    // Append row first to get the row number, then set formula
    sheet.appendRow([id, date, startTime, endTime, "", note || ""]);
    const lastRow = sheet.getLastRow();
    const formula = `=TEXT(MOD(D${lastRow}-C${lastRow},1)*24,"0.00")`;
    sheet.getRange(lastRow, 5).setFormula(formula);

    SpreadsheetApp.flush();
    const hours = sheet.getRange(lastRow, 5).getDisplayValue();

    return { success: true, totalHours: hours };
}

/**
 * Returns app configuration.
 */
function getAppConfig() {
    return {
        cycleStartDay: getCycleStartDay()
    };
}

/**
 * Updates app configuration.
 * @param {Object} data - { cycleStartDay: number }
 */
function updateAppConfig(data) {
    const { cycleStartDay } = data;
    const day = parseInt(cycleStartDay, 10);
    if (isNaN(day) || day < 1 || day > 28) {
        throw new Error("Cycle start day must be between 1 and 28");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('_Config');
    if (!sheet) {
        sheet = ss.insertSheet('_Config');
        sheet.getRange('A1').setValue('CycleStartDay');
    }
    sheet.getRange('B1').setValue(day);

    // Update the global variable
    CYCLE_START_DAY = day;

    return { success: true, cycleStartDay: day };
}
