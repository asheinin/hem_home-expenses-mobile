/**
 * Mobile Add Expense Web App — Home Expenses
 *
 * Standalone GAS web app (separate from the HomeExpenses script).
 * Accesses the Home Expenses spreadsheet via its ID, stored in Script Properties.
 *
 * Setup (one-time):
 *   1. In the Apps Script editor for THIS project, go to
 *      Project Settings → Script Properties and add:
 *        SPREADSHEET_ID = <the ID from the Home Expenses spreadsheet URL>
 *   2. Deploy as Web App:
 *        Execute as: Me (owner)
 *        Who has access: Anyone with Google Account
 *
 * Settings page: append ?page=settings to the web app URL.
 */

// ============================================================================
// Constants
// ============================================================================

var SPREADSHEET_NAME_PREFIX = 'Home Expenses'; // used when searching by name
var AUTO_SWITCH_TRIGGER_HANDLER = 'autoSwitchToNewYearFile';

// ============================================================================
// Helpers
// ============================================================================

function _getSpreadsheetId() {
    var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!id) throw new Error('SPREADSHEET_ID is not set. Open the Settings page to configure it.');
    return id;
}

function _getSpreadsheet() {
    return SpreadsheetApp.openById(_getSpreadsheetId());
}

function _monthName(monthNum) {
    var months = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ];
    return months[monthNum - 1] || String(monthNum);
}

/** Extract a spreadsheet ID from a Google Sheets URL, or return as-is if already an ID. */
function _extractId(input) {
    input = input.trim();
    // Match /d/<id>/ or /spreadsheets/d/<id>
    var m = input.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
    if (m) return m[1];
    // Looks like a raw ID (no slashes)
    if (/^[a-zA-Z0-9_-]{25,}$/.test(input)) return input;
    return null;
}

/** Find the most recent "Home Expenses {year}" file in Drive and return {id, name}. */
function _findFileByName(name) {
    var files = DriveApp.getFilesByName(name);
    if (files.hasNext()) {
        var f = files.next();
        return { id: f.getId(), name: f.getName() };
    }
    return null;
}

/** Check whether the yearly auto-switch trigger is installed. */
function _autoSwitchTriggerInstalled() {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === AUTO_SWITCH_TRIGGER_HANDLER) return true;
    }
    return false;
}

/** 
 * Generate a dynamic PWA manifest and encode it as Base64 to inject into HTML directly.
 * This bypasses Google Apps Script's cross-domain web app iframe restrictions.
 */
function _getBase64Manifest() {
    var manifest = {
        "short_name": "HE",
        "name": "HomeExpenses",
        "icons": [
            {
                "src": "https://drive.google.com/uc?export=view&id=1GyoYGc2ldx_mCHklAEjFVDG0wrMeQODD",
                "type": "image/png",
                "sizes": "192x192"
            },
            {
                "src": "https://drive.google.com/uc?export=view&id=1GyoYGc2ldx_mCHklAEjFVDG0wrMeQODD",
                "type": "image/png",
                "sizes": "512x512"
            }
        ],
        "start_url": ScriptApp.getService().getUrl(),
        "display": "standalone",
        "theme_color": "#667eea",
        "background_color": "#f8f9fc"
    };
    return Utilities.base64Encode(JSON.stringify(manifest));
}

// ============================================================================
// Web App entry point
// ============================================================================

function doGet(e) {
    var page = (e && e.parameter && e.parameter.page) || 'main';

    if (page === 'settings') {
        return _serveSettings();
    }
    return _serveMain();
}

function _serveMain() {
    var myNumbers = new staticNumbers();
    var ss;

    try {
        ss = _getSpreadsheet();
    } catch (err) {
        // Spreadsheet not configured yet — redirect to settings
        var errTemplate = HtmlService.createTemplateFromFile('ui/MobileAddExpense');
        errTemplate.expenseNames = [];
        errTemplate.expenseTypes = [];
        errTemplate.expensePeriods = [];
        errTemplate.spouseNames = [];
        errTemplate.configError = err.message;
        return errTemplate
            .evaluate()
            .setTitle('Add Expense · Home Expenses')
            .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, shrink-to-fit=no')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    var sheets = ss.getSheets();
    var expenseNames = new Set();
    var expenseTypes = new Set();
    var expensePeriods = new Set();
    var nameCategoryMap = {};
    
    var currentMonthIdx = new Date().getMonth() + 1;

    for (var i = 1; i < Math.min(sheets.length, 12); i++) {
        var sheet = sheets[i];
        var numRows = myNumbers.expenseLastRow - myNumbers.expenseFirstRow + 1;
        var values = sheet
            .getRange(myNumbers.expenseFirstRow, 1, numRows, myNumbers.expensePAPColumn)
            .getValues();

        values.forEach(function (row) {
            var n = row[myNumbers.expenseDescrColumn - 1];
            var t = row[myNumbers.expenseTypeColumn - 1];
            var p = row[myNumbers.expencePeriodColumn - 1];
            
            if (n && i === currentMonthIdx) {
                var nStr = n.toString().trim();
                expenseNames.add(nStr);
                
                // Store fastest category mapping for synchronous client resolution
                var tStr = t ? t.toString().trim() : '';
                if (tStr && !nameCategoryMap[nStr.toLowerCase()]) {
                    nameCategoryMap[nStr.toLowerCase()] = tStr;
                }
            }
            if (t) expenseTypes.add(t.toString().trim());
            if (p) expensePeriods.add(p.toString().trim());
        });
    }

    var dashboard = sheets[0];
    var spouse1 = dashboard.getRange(myNumbers.dashNamesRow, myNumbers.dashSpouse1NameColumn).getValue();
    var spouse2 = dashboard.getRange(myNumbers.dashNamesRow, myNumbers.dashSpouse2NameColumn).getValue();

    var template = HtmlService.createTemplateFromFile('ui/MobileAddExpense');
    template.expenseNames = Array.from(expenseNames).sort(function (a, b) {
        return a.toLowerCase().localeCompare(b.toLowerCase());
    });
    template.expenseTypes = Array.from(expenseTypes).sort(function (a, b) {
        return a.toLowerCase().localeCompare(b.toLowerCase());
    });
    template.expensePeriods = Array.from(expensePeriods).sort();
    template.nameCategoryMap = nameCategoryMap;
    template.spouseNames = [spouse1, spouse2].filter(Boolean);
    var now = new Date();
    template.currentDateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM dd, yyyy");
    template.configError = null;
    template.spreadsheetName = ss.getName();
    template.manifestBase64 = _getBase64Manifest();

    return template
        .evaluate()
        .setTitle('Add Expense · Home Expenses')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, shrink-to-fit=no')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _serveSettings() {
    var props = PropertiesService.getScriptProperties();
    var currentId = props.getProperty('SPREADSHEET_ID') || '';
    var currentName = '';

    if (currentId) {
        try {
            currentName = SpreadsheetApp.openById(currentId).getName();
        } catch (e) {
            currentName = '(file not accessible — ID may be stale)';
        }
    }

    var template = HtmlService.createTemplateFromFile('ui/Settings');
    template.currentId = currentId;
    template.currentName = currentName;
    template.autoSwitchInstalled = _autoSwitchTriggerInstalled();
    template.nextYear = new Date().getFullYear() + 1;
    template.manifestBase64 = _getBase64Manifest();

    return template
        .evaluate()
        .setTitle('Settings · Home Expenses Mobile')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, shrink-to-fit=no')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================================
// Settings actions (called from Settings page via google.script.run)
// ============================================================================

/**
 * Save spreadsheet configuration.
 * @param {string} input  Google Sheets URL, spreadsheet ID, or Drive file name
 * @returns {{success: boolean, message: string, name: string}}
 */
function saveSpreadsheetConfig(input) {
    try {
        input = (input || '').trim();
        if (!input) return { success: false, message: 'Please enter a URL, ID, or file name.' };

        var id = _extractId(input);
        var name = '';

        if (id) {
            // URL or raw ID
            try {
                name = SpreadsheetApp.openById(id).getName();
            } catch (e) {
                return { success: false, message: 'Could not open spreadsheet. Check the URL/ID and make sure it is shared with this script\'s owner account.' };
            }
        } else {
            // Treat as a file name — search Drive
            var found = _findFileByName(input);
            if (!found) {
                return { success: false, message: 'File "' + input + '" not found in Google Drive. Make sure the name matches exactly.' };
            }
            id = found.id;
            name = found.name;
        }

        PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
        return { success: true, message: 'Connected to "' + name + '".', name: name };

    } catch (err) {
        Logger.log(err);
        return { success: false, message: 'Unexpected error: ' + err.toString() };
    }
}

/**
 * Install or remove the January 1st auto-switch trigger.
 * @param {boolean} install  true = install, false = remove
 * @returns {{success: boolean, message: string, installed: boolean}}
 */
function setAutoSwitchTrigger(install) {
    try {
        // Remove existing trigger first (avoid duplicates)
        ScriptApp.getProjectTriggers().forEach(function (t) {
            if (t.getHandlerFunction() === AUTO_SWITCH_TRIGGER_HANDLER) ScriptApp.deleteTrigger(t);
        });

        if (install) {
            var triggerBuilder = ScriptApp.newTrigger(AUTO_SWITCH_TRIGGER_HANDLER).timeBased();

            triggerBuilder
                .onMonthDay(1)
                .atHour(6)   // Runs on the 1st of every month at 6 AM
                .create();
            return { success: true, installed: true, message: 'Auto-switch enabled. On January 1st the app will automatically find and connect to the new Home Expenses file.' };
        } else {
            return { success: true, installed: false, message: 'Auto-switch disabled.' };
        }
    } catch (err) {
        Logger.log(err);
        return { success: false, installed: _autoSwitchTriggerInstalled(), message: 'Error: ' + err.toString() };
    }
}

/**
 * Get current settings (called on page load to refresh state without reload).
 * @returns {{id: string, name: string, autoSwitchInstalled: boolean}}
 */
function getSettings() {
    var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '';
    var name = '';
    if (id) {
        try { name = SpreadsheetApp.openById(id).getName(); } catch (e) { name = '(not accessible)'; }
    }
    return { id: id, name: name, autoSwitchInstalled: _autoSwitchTriggerInstalled() };
}

// ============================================================================
// Auto year-switch (triggered on Jan 1)
// ============================================================================

/**
 * Time-based trigger handler: runs on January 1st each year.
 * Searches Google Drive for "Home Expenses {currentYear}" and updates SPREADSHEET_ID.
 */
function autoSwitchToNewYearFile() {
    var now = new Date();
    // Only proceed if it is January (getMonth() returns 0 for January)
    if (now.getMonth() !== 0) return;

    var year = now.getFullYear();
    var targetName = SPREADSHEET_NAME_PREFIX + ' ' + year;

    Logger.log('Auto year-switch: looking for "' + targetName + '"');

    var found = _findFileByName(targetName);
    if (!found) {
        Logger.log('Auto year-switch: file not found — "' + targetName + '"');
        return;
    }

    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', found.id);
    Logger.log('Auto year-switch: switched to "' + found.name + '" (' + found.id + ')');
}

// ============================================================================
// Current month expense reference (called from client)
// ============================================================================

/**
 * Returns capacity info for the current month's expense range.
 * Used to display "X rows left" and block submission when the sheet is full.
 * @returns {{success: boolean, totalSlots: number, usedSlots: number, emptySlots: number, isFull: boolean}}
 */
function getMonthCapacity() {
    try {
        var myNumbers = new staticNumbers();
        var ss = _getSpreadsheet();
        var currentMonth = new Date().getMonth() + 1; // 1-based
        var sheet = ss.getSheets()[currentMonth];
        var numRows = myNumbers.expenseLastRow - myNumbers.expenseFirstRow + 1;
        var values = sheet
            .getRange(myNumbers.expenseFirstRow, 1, numRows, myNumbers.expenseAmountColumn)
            .getValues();

        var descrIdx = myNumbers.expenseDescrColumn - 1;
        var amtIdx   = myNumbers.expenseAmountColumn - 1;

        var usedSlots  = 0;
        var emptySlots = 0;

        values.forEach(function (row) {
            var hasDescr = row[descrIdx] !== '' && row[descrIdx] != null;
            var hasAmt   = row[amtIdx]   !== '' && row[amtIdx]   != null;
            if (hasDescr || hasAmt) {
                usedSlots++;
            } else {
                emptySlots++;
            }
        });

        return {
            success:    true,
            totalSlots: numRows,
            usedSlots:  usedSlots,
            emptySlots: emptySlots,
            isFull:     emptySlots === 0
        };
    } catch (err) {
        Logger.log(err);
        return { success: false, totalSlots: 0, usedSlots: 0, emptySlots: 0, isFull: false };
    }
}

/**
 * Returns all expenses for the current month.
 * @returns {{success: boolean, expenses: Array<{name:string, type:string, amount:string}>, month: string}}
 */
function getCurrentMonthExpenses() {
    try {
        var myNumbers = new staticNumbers();
        var ss = _getSpreadsheet();
        var date = new Date();
        var currentMonth = date.getMonth() + 1; // 1-based
        var sheet = ss.getSheets()[currentMonth]; // sheets[1] = January, etc.
        var numRows = myNumbers.expenseLastRow - myNumbers.expenseFirstRow + 1;
        var values = sheet
            .getRange(myNumbers.expenseFirstRow, 1, numRows, myNumbers.expensePAPColumn)
            .getValues();

        var expenses = [];
        values.forEach(function (row) {
            var name   = (row[myNumbers.expenseDescrColumn  - 1] || '').toString().trim();
            var type   = (row[myNumbers.expenseTypeColumn   - 1] || '').toString().trim();
            var amount = row[myNumbers.expenseAmountColumn  - 1];
            var split  = (row[myNumbers.expenceSplitColumn  - 1] || '').toString().trim() === 'Y';
            var pap    = (row[myNumbers.expensePAPColumn    - 1] || '').toString().trim() === 'PAP';
            if (name) {
                var amtStr = (amount !== '' && amount !== null && !isNaN(parseFloat(amount)))
                    ? '$' + parseFloat(amount).toFixed(2)
                    : '—';
                expenses.push({ name: name, type: type, amount: amtStr, split: split, pap: pap });
            }
        });

        // Sort by type, then name
        expenses.sort(function (a, b) {
            var t = a.type.localeCompare(b.type);
            return t !== 0 ? t : a.name.localeCompare(b.name);
        });

        return { success: true, expenses: expenses, month: _monthName(currentMonth) };
    } catch (err) {
        Logger.log(err);
        return { success: false, expenses: [], month: '', message: err.toString() };
    }
}

// ============================================================================
// Expense form submission (mobile-safe — no getUi() calls)
// ============================================================================

/**
 * Called from the client via google.script.run.mobileProcessForm(...)
 *
 * @param {string}  newExpenseItem  Expense description / name
 * @param {string}  newExpenseType  Category / type
 * @param {string}  expenseAmount   Dollar amount as string; empty = unpopulated
 * @param {boolean} pap             Pre-authorized payment flag
 * @param {string}  expensePeriod   Billing period (e.g. "Monthly")
 * @param {string}  mode            'ot' = one-time this month | 'rm' = recurrent from this month onward
 * @param {boolean} split           Split between spouses flag
 * @param {boolean} paid            Paid flag
 * @param {string}  paidByName      Display name of paying spouse
 * @param {number}  paidByIndex     0 = not specified, 1 = spouse2 pays, 2 = spouse1 pays
 * @returns {{success: boolean, message: string}}
 */
function mobileProcessForm(
    newExpenseItem,
    newExpenseType,
    expenseAmount,
    pap,
    expensePeriod,
    mode,
    split,
    paid,
    paidByName,
    paidByIndex,
    actionType
) {
    try {
        var myNumbers = new staticNumbers();

        if (!newExpenseItem) return { success: false, message: 'Expense Name is required.' };
        if (!newExpenseType) return { success: false, message: 'Category is required.' };

        var amount;
        if (expenseAmount !== '' && expenseAmount !== null && expenseAmount !== undefined) {
            amount = parseFloat(expenseAmount);
            if (isNaN(amount)) {
                return { success: false, message: 'Amount must be a valid number.' };
            }
            amount = parseFloat(amount.toFixed(2));
        } else {
            amount = -1;
        }

        var date = new Date();
        var currentMonth = date.getMonth() + 1;
        var mStart = currentMonth;
        var mEnd = (mode === 'rm') ? 12 : currentMonth;

        var ss = _getSpreadsheet();
        var numOfRows = myNumbers.expenseLastRow - myNumbers.expenseFirstRow + 1;
        var amountColIdx = myNumbers.expenseAmountColumn - 1;

        for (var i = mStart; i <= mEnd; i++) {
            var sheet = ss.getSheets()[i];
            var descriptions = sheet
                .getRange(myNumbers.expenseFirstRow, myNumbers.expenseDescrColumn, numOfRows)
                .getValues()
                .flat();
            
            var types = sheet
                .getRange(myNumbers.expenseFirstRow, myNumbers.expenseTypeColumn, numOfRows)
                .getValues()
                .flat();

            var existingIndex = -1;
            var isUpdateSimilar = actionType && actionType.indexOf('_update_similar|') === 0;
            var isAddToExisting = (actionType === '_add_to_existing');
            var targetNameForUpdate = isUpdateSimilar ? actionType.split('|')[1] : null;

            if (actionType !== '_add_new') {
                if (actionType === '_update' || isAddToExisting) {
                    // 1. Exact name + type match
                    for (var k = 0; k < descriptions.length; k++) {
                        var dStr = (descriptions[k] || '').toString().trim().toLowerCase();
                        var tStr = (types[k] || '').toString().trim().toLowerCase();
                        if (dStr === newExpenseItem.toLowerCase().trim() && tStr === newExpenseType.toLowerCase().trim()) {
                            existingIndex = k;
                            break;
                        }
                    }
                    // 2. Name-only match (capacity-full fallback)
                    if (existingIndex === -1 && isAddToExisting) {
                        for (var k = 0; k < descriptions.length; k++) {
                            var dStr = (descriptions[k] || '').toString().trim().toLowerCase();
                            if (dStr === newExpenseItem.toLowerCase().trim()) {
                                existingIndex = k;
                                break;
                            }
                        }
                    }
                    // 3. Type-only match (capacity-full fallback)
                    if (existingIndex === -1 && isAddToExisting) {
                        for (var k = 0; k < descriptions.length; k++) {
                            var tStr = (types[k] || '').toString().trim().toLowerCase();
                            if (tStr === newExpenseType.toLowerCase().trim()) {
                                existingIndex = k;
                                break;
                            }
                        }
                    }
                    // 4. Last occupied row (capacity-full last resort)
                    if (existingIndex === -1 && isAddToExisting) {
                        for (var k = descriptions.length - 1; k >= 0; k--) {
                            if ((descriptions[k] || '').toString().trim() !== '') {
                                existingIndex = k;
                                break;
                            }
                        }
                    }
                } else if (isUpdateSimilar) {
                    for (var k = 0; k < descriptions.length; k++) {
                        var dStr = (descriptions[k] || '').toString().trim().toLowerCase();
                        var tStr = (types[k] || '').toString().trim().toLowerCase();
                        if (dStr === targetNameForUpdate.toLowerCase().trim() && tStr === newExpenseType.toLowerCase().trim()) {
                            existingIndex = k;
                            break;
                        }
                    }
                } else {
                    for (var k = 0; k < descriptions.length; k++) {
                        var dStr = (descriptions[k] || '').toString().trim().toLowerCase();
                        if (dStr === newExpenseItem.toLowerCase().trim()) {
                            existingIndex = k;
                            break;
                        }
                    }
                }
            }

            if (existingIndex !== -1) {
                var row = existingIndex + myNumbers.expenseFirstRow;

                if (actionType === '_update' || isUpdateSimilar || isAddToExisting) {
                    // Accumulate amount
                    var currentAmt = sheet.getRange(row, myNumbers.expenseAmountColumn).getValue();
                    var newTotal = (parseFloat(currentAmt) || 0) + (amount !== -1 ? amount : 0);
                    sheet.getRange(row, myNumbers.expenseAmountColumn).setValue(newTotal);

                    if (isUpdateSimilar) {
                        var currentName = sheet.getRange(row, myNumbers.expenseDescrColumn).getValue();
                        sheet.getRange(row, myNumbers.expenseDescrColumn).setValue(currentName + ', ' + newExpenseItem);
                    }

                    // Accumulate who-pays columns
                    if (amount !== -1) {
                        if (paidByIndex == 1) {
                            var p1 = sheet.getRange(row, myNumbers.expenseSecondPayColumn).getValue();
                            sheet.getRange(row, myNumbers.expenseSecondPayColumn).setValue((parseFloat(p1) || 0) + amount);
                        } else if (paidByIndex == 2) {
                            var p2 = sheet.getRange(row, myNumbers.expenseFirstPayColumn).getValue();
                            sheet.getRange(row, myNumbers.expenseFirstPayColumn).setValue((parseFloat(p2) || 0) + amount);
                        }
                    }

                    sheet.getRange(row, myNumbers.expensePAPColumn).setValue(pap ? 'PAP' : '');
                    sheet.getRange(row, myNumbers.expencePeriodColumn).setValue(expensePeriod);
                    sheet.getRange(row, myNumbers.expenceSplitColumn).setValue(split ? 'Y' : 'N');
                    sheet.getRange(row, myNumbers.expensePaidColumn).setValue(paid ? 'Y' : '');
                
                } else {
                    sheet.getRange(row, myNumbers.expenseFirstPayColumn, 1, 2).clearContent();

                    if (amount !== -1) {
                        sheet.getRange(row, myNumbers.expenseAmountColumn).setValue(amount);
                        if (paidByIndex == 1) sheet.getRange(row, myNumbers.expenseSecondPayColumn).setValue(amount);
                        if (paidByIndex == 2) sheet.getRange(row, myNumbers.expenseFirstPayColumn).setValue(amount);
                    }

                    sheet.getRange(row, myNumbers.expenseTypeColumn).setValue(newExpenseType);
                    sheet.getRange(row, myNumbers.expensePAPColumn).setValue(pap ? 'PAP' : '');
                    sheet.getRange(row, myNumbers.expencePeriodColumn).setValue(expensePeriod);
                    sheet.getRange(row, myNumbers.expenceSplitColumn).setValue(split ? 'Y' : 'N');
                    sheet.getRange(row, myNumbers.expensePaidColumn).setValue(paid ? 'Y' : '');
                }

            } else {
                var expenseData = sheet
                    .getRange(myNumbers.expenseFirstRow, 1, numOfRows, myNumbers.expenseAmountColumn)
                    .getValues();
                var inserted = false;

                for (var j = 0; j < numOfRows; j++) {
                    var rowData = expenseData[j];
                    if ((rowData[0] === '' || rowData[0] == null) &&
                        (rowData[amountColIdx] === '' || rowData[amountColIdx] == null)) {
                        var newRow = j + myNumbers.expenseFirstRow;

                        sheet.getRange(newRow, myNumbers.expenseDescrColumn).setValue(newExpenseItem);
                        sheet.getRange(newRow, myNumbers.expenseTypeColumn).setValue(newExpenseType);

                        if (amount !== -1) {
                            sheet.getRange(newRow, myNumbers.expenseAmountColumn).setValue(amount);
                            if (paidByIndex == 1) sheet.getRange(newRow, myNumbers.expenseSecondPayColumn).setValue(amount);
                            if (paidByIndex == 2) sheet.getRange(newRow, myNumbers.expenseFirstPayColumn).setValue(amount);
                        }

                        sheet.getRange(newRow, myNumbers.expensePAPColumn).setValue(pap ? 'PAP' : '');
                        sheet.getRange(newRow, myNumbers.expencePeriodColumn).setValue(expensePeriod);
                        sheet.getRange(newRow, myNumbers.expenceSplitColumn).setValue(split ? 'Y' : 'N');
                        sheet.getRange(newRow, myNumbers.expensePaidColumn).setValue(paid ? 'Y' : '');

                        inserted = true;
                        break;
                    }
                }

                if (!inserted) {
                    return { success: false, message: 'No space to add expense in ' + _monthName(i) + '.' };
                }
            }
        }

        var label = mode === 'rm'
            ? 'from ' + _monthName(mStart) + ' through December'
            : 'in ' + _monthName(mStart);
        return { success: true, message: '"' + newExpenseItem + '" added ' + label + '.' };

    } catch (err) {
        Logger.log(err);
        return { success: false, message: 'Error: ' + err.toString() };
    }
}
