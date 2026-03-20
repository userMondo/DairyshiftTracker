/**
 * Migrate.gs — Run fixCycleDay() to re-sort rows based on new CYCLE_START_DAY = 9.
 *
 * Rows dated on the 8th of any month will be moved from the current cycle sheet
 * to the previous month's cycle sheet.
 *
 * Also re-applies Total Hours formula to all rows.
 *
 * After running, delete this file from your GAS project.
 */
function fixCycleDay() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cycleDay = CYCLE_START_DAY; // from Config.gs (should be 9)
  Logger.log('Fixing cycle with CYCLE_START_DAY = ' + cycleDay);

  // Find all cycle sheets (e.g., "1/26", "2/26", "3/26")
  const cycleSheets = [];
  for (const sheet of ss.getSheets()) {
    if (sheet.getName().match(/^\d{1,2}\/\d{2}$/)) {
      cycleSheets.push(sheet);
    }
  }

  if (cycleSheets.length === 0) {
    Logger.log('No cycle sheets found.');
    return;
  }
  Logger.log('Found sheets: ' + cycleSheets.map(function(s) { return s.getName(); }).join(', '));

  // Step 1: Collect ALL rows from all cycle sheets
  var allRows = [];
  for (const sheet of cycleSheets) {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue;

    const numCols = Math.min(Math.max(sheet.getLastColumn(), 6), 6);
    const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

    for (const row of data) {
      if (!row[1]) continue; // Skip empty rows
      // Pad to 6 columns
      while (row.length < 6) row.push('');
      allRows.push(row);
    }
  }

  Logger.log('Total rows collected: ' + allRows.length);

  // Step 2: Group rows by correct cycle sheet based on new CYCLE_START_DAY
  var cycleData = {};
  for (var i = 0; i < allRows.length; i++) {
    var row = allRows[i];
    var dateVal = row[1];
    var day, month, year;

    if (dateVal instanceof Date) {
      day = dateVal.getDate();
      month = dateVal.getMonth() + 1;
      year = dateVal.getFullYear();
    } else {
      var parts = String(dateVal).split('-');
      if (parts.length !== 3) continue;
      day = parseInt(parts[2], 10);
      month = parseInt(parts[1], 10);
      year = parseInt(parts[0], 10);
    }

    if (day < cycleDay) {
      month -= 1;
      if (month < 1) { month = 12; year -= 1; }
    }

    var targetName = month + '/' + String(year).slice(-2);
    if (!cycleData[targetName]) cycleData[targetName] = [];

    // Clear column E (will be replaced by formula)
    row[4] = '';
    cycleData[targetName].push(row);
  }

  // Step 3: Delete all old cycle sheets
  for (const sheet of cycleSheets) {
    ss.deleteSheet(sheet);
  }

  // Step 4: Create new sheets and write data
  var formulaCount = 0;
  for (var sheetName in cycleData) {
    var rows = cycleData[sheetName];

    // Sort by date ascending
    rows.sort(function(a, b) {
      var da = a[1] instanceof Date ? a[1].getTime() : new Date(a[1]).getTime();
      var db = b[1] instanceof Date ? b[1].getTime() : new Date(b[1]).getTime();
      return da - db;
    });

    var sheet = ss.insertSheet(sheetName, 0);
    sheet.appendRow(['ID', 'Date', 'Start Time', 'End Time', 'Total Hours', 'Note']);
    sheet.setFrozenRows(1);

    // Write rows
    var startRow = 2;
    sheet.getRange(startRow, 1, rows.length, 6).setValues(rows);

    // Format columns: Date as dd/MM/yy, Start/End Time as HH:mm:ss
    sheet.getRange(startRow, 2, rows.length, 1).setNumberFormat('dd/MM/yy');
    sheet.getRange(startRow, 3, rows.length, 1).setNumberFormat('HH:mm:ss');
    sheet.getRange(startRow, 4, rows.length, 1).setNumberFormat('HH:mm:ss');
    SpreadsheetApp.flush();

    // Set Total Hours formula for each row
    for (var r = 0; r < rows.length; r++) {
      var rowNum = startRow + r;
      if (rows[r][3]) { // End Time exists
        var formula = '=TEXT(MOD(D' + rowNum + '-C' + rowNum + ',1)*24,"0.00")';
        sheet.getRange(rowNum, 5).setFormula(formula);
        formulaCount++;
      }
    }
    SpreadsheetApp.flush();
  }

  Logger.log('Done! Created sheets: ' + Object.keys(cycleData).join(', '));
  Logger.log('Applied formula to ' + formulaCount + ' rows.');
  Logger.log('You can now delete Migrate.gs from your project.');
}
