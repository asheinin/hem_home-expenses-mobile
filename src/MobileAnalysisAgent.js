/**
 * Mobile Expense Analysis Agent
 *
 * Provides two server-side entry points for the mobile web app:
 *  - getMonthlySummary()        → current-month running total + who owes
 *  - runMobileExpenseAnalysis() → full AI agent (comparison, forecast, spikes, Gemini)
 *
 * Ported from HomeExpenses/src/ai/ExpenseAnalysisAgent.js and adjusted
 * to use _getSpreadsheet() (via Script Properties SPREADSHEET_ID) instead of
 * SpreadsheetApp.getActiveSpreadsheet().
 */

// ============================================================================
// Monthly Summary (running total + who owes) — shown on top of main page
// ============================================================================

/**
 * Returns the current-month running total and who-owes information
 * from the Dashboard sheet.
 *
 * @returns {{
 *   success: boolean,
 *   month: string,
 *   runningTotal: number,
 *   sp1Name: string,
 *   sp2Name: string,
 *   sp1Balance: number,
 *   sp2Balance: number,
 *   oweeName: string,
 *   oweeAmount: number
 * }}
 */
function getMonthlySummary() {
    try {
        var myNumbers = new staticNumbers();
        var ss = _getSpreadsheet();
        var dashboard = ss.getSheets()[0];
        var now = new Date();
        var currentMonthIdx = now.getMonth() + 1; // 1-based

        // Spouse names — batch read both columns in one call
        var namesRow = dashboard.getRange(myNumbers.dashNamesRow, myNumbers.dashSpouse1NameColumn, 1, myNumbers.dashSpouse2NameColumn - myNumbers.dashSpouse1NameColumn + 1).getValues()[0];
        var sp1Name = (namesRow[0] || '').toString().trim();
        var sp2Name = (namesRow[namesRow.length - 1] || '').toString().trim();

        // Read the month total from the SUM cell at expenseLastRow+1 in the expense sheet.
        // This cell contains the spreadsheet SUM formula and correctly handles positive and negative amounts.
        var monthSheet = ss.getSheets()[currentMonthIdx];
        var runningTotal = 0;
        if (monthSheet) {
            var totalCell = monthSheet.getRange(myNumbers.expenseTotalAmountRowBeforeSplit, myNumbers.expenseAmountColumn);
            var cellValue = totalCell.getValue();
            runningTotal = parseFloat(cellValue) || 0;
        }

        // Read balance columns from Dashboard for the current month row
        // Dashboard row = dashFirstMonthRow + (monthIndex - 1)  [rows 7-18 for Jan-Dec]
        var dashRowIndex = myNumbers.dashFirstMonthRow + (currentMonthIdx - 1);
        var dashData = dashboard
            .getRange(dashRowIndex, 1, 1, myNumbers.dashColumns)
            .getValues()[0];

        var sp1Balance = parseFloat(dashData[myNumbers.dashSp1BalanceColumn - 1]) || 0;
        var sp2Balance = parseFloat(dashData[myNumbers.dashSp2BalanceColumn - 1]) || 0;

        // Positive balance means that spouse owes
        var oweeName = '';
        var oweeAmount = 0;
        if (sp1Balance > 0.5) {
            oweeName = sp1Name;
            oweeAmount = sp1Balance;
        } else if (sp2Balance > 0.5) {
            oweeName = sp2Name;
            oweeAmount = sp2Balance;
        }

        return {
            success: true,
            month: _monthName(currentMonthIdx),
            runningTotal: runningTotal,
            sp1Name: sp1Name,
            sp2Name: sp2Name,
            sp1Balance: sp1Balance,
            sp2Balance: sp2Balance,
            oweeName: oweeName,
            oweeAmount: oweeAmount
        };
    } catch (err) {
        Logger.log('getMonthlySummary error: ' + err);
        return { success: false, message: err.toString() };
    }
}


// ============================================================================
// Expense Analysis Agent — full AI analysis
// ============================================================================

/**
 * Main entry point called from the mobile client via google.script.run.
 * Returns a JSON-serialisable result object suitable for rendering.
 *
 * @returns {{success: boolean, results: Object, message: string}}
 */
function runMobileExpenseAnalysis() {
    try {
        var myNumbers = new staticNumbers();
        var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        var now = new Date();
        var currentMonthIndex = now.getMonth();
        var currentMonthName = months[currentMonthIndex];
        var currentYear = now.getFullYear();

        var prevMonthDate = new Date(currentYear, currentMonthIndex - 1, 1);
        var prevMonthIndex = prevMonthDate.getMonth();
        var prevMonthName = months[prevMonthIndex];
        var prevMonthYear = prevMonthDate.getFullYear();

        var ss = _getSpreadsheet();

        var comparisonData = _mobileGetMonthlyComparisonData(
            ss, currentMonthIndex, currentYear, prevMonthIndex, prevMonthYear, myNumbers, months
        );

        var forecastData = _mobileCalculateAnnualForecast(
            ss, currentMonthIndex, currentYear, myNumbers, months
        );

        var spikeAnalysis = _mobileDetectExpenseSpikes(
            comparisonData.current, comparisonData.yearAgo, myNumbers
        );

        var useAIToggleStr = PropertiesService.getScriptProperties().getProperty('USE_AI_TOGGLE');
        var aiEnabled = useAIToggleStr === null ? true : (useAIToggleStr === 'true');

        // Gemini call is deferred — client fetches it via runMobileAIInsights() so the
        // analytics sheet can render charts/forecast/spikes immediately (~3-8s saved
        // perceptually).

        // ── YTD category totals (sum Jan → current month) ────────────────────
        var ytdCategoryTotals = {};
        var currentSS = _mobileGetSpreadsheetForYear(currentYear);
        if (currentSS) {
            for (var m = 0; m <= currentMonthIndex; m++) {
                var mStats = _mobileGetMonthStats(currentSS, months[m], currentYear, myNumbers);
                if (mStats && mStats.categoryTotals) {
                    for (var cat in mStats.categoryTotals) {
                        ytdCategoryTotals[cat] = (ytdCategoryTotals[cat] || 0) + mStats.categoryTotals[cat];
                    }
                }
            }
        }

        var results = {
            timestamp: now.toISOString(),
            currentMonth: currentMonthName,
            currentYear: currentYear,
            comparison: comparisonData,
            forecast: forecastData,
            spikes: spikeAnalysis,
            aiInsights: null,
            aiPending: aiEnabled,
            currentMonthCategoryTotals: (comparisonData.current && comparisonData.current.categoryTotals) || {},
            ytdCategoryTotals: ytdCategoryTotals,
            aiEnabled: aiEnabled
        };

        return { success: true, results: results };
    } catch (err) {
        Logger.log('runMobileExpenseAnalysis error: ' + err);
        return { success: false, message: err.toString() };
    }
}

/**
 * Generate Gemini AI insights from a previously computed core analysis payload.
 * Called from the client immediately after runMobileExpenseAnalysis() returns,
 * so the analytics sheet can paint without waiting for the LLM round-trip.
 *
 * @param {{comparison: Object, forecast: Object, spikes: Object}} payload
 * @returns {{success: boolean, aiInsights?: string, message?: string}}
 */
function runMobileAIInsights(payload) {
    try {
        if (!payload) return { success: false, message: 'No analysis payload provided.' };

        var useAIToggleStr = PropertiesService.getScriptProperties().getProperty('USE_AI_TOGGLE');
        var aiEnabled = useAIToggleStr === null ? true : (useAIToggleStr === 'true');
        if (!aiEnabled) return { success: false, message: 'AI Insights are disabled in settings.' };

        var insights = _mobileGenerateAgentAnalysis(
            payload.comparison || {},
            payload.forecast || {},
            payload.spikes || { hasAnomalies: false, spikes: [], aboveNormal: [] }
        );
        return { success: true, aiInsights: insights };
    } catch (err) {
        Logger.log('runMobileAIInsights error: ' + err);
        return { success: false, message: err.toString() };
    }
}



// ============================================================================
// Internal helpers (mobile-adapted versions of agent helpers)
// ============================================================================

/**
 * Per-execution cache: avoids re-opening the same spreadsheet multiple times
 * within a single server-side call. GAS execution context is ephemeral so
 * this object is reset on every new invocation automatically.
 */
var _mobileSSCache = {};

function _mobileGetSpreadsheetForYear(year) {
    if (_mobileSSCache[year]) return _mobileSSCache[year];

    var ss = _getSpreadsheet();
    if (ss.getName().indexOf(year.toString()) !== -1) {
        _mobileSSCache[year] = ss;
        return ss;
    }

    // Search Drive for "Home Expenses {YYYY}" (e.g. "Home Expenses 2025")
    var targetName = 'Home Expenses ' + year;
    var files = DriveApp.getFilesByName(targetName);

    while (files.hasNext()) {
        var file = files.next();
        if (!file.isTrashed()) {
            var found = SpreadsheetApp.openById(file.getId());
            _mobileSSCache[year] = found;
            return found;
        }
    }

    // In case there are older files named "Home payments {YYYY}"
    var oldTargetName = 'Home payments ' + year;
    var oldFiles = DriveApp.getFilesByName(oldTargetName);

    while (oldFiles.hasNext()) {
        var oldFile = oldFiles.next();
        if (!oldFile.isTrashed()) {
            var oldFound = SpreadsheetApp.openById(oldFile.getId());
            _mobileSSCache[year] = oldFound;
            return oldFound;
        }
    }

    _mobileSSCache[year] = null;
    return null;
}

function _mobileNormalizeExpenseDay(dateVal, monthIndex, year, daysInMonth) {
    if (!dateVal) return 1;
    var d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(d.getTime())) return 1;

    var firstOfMonth = new Date(year, monthIndex, 1, 0, 0, 0, 0);
    var lastOfMonth = new Date(year, monthIndex, daysInMonth, 23, 59, 59, 999);

    if (d < firstOfMonth) return 1;
    if (d > lastOfMonth) return daysInMonth;
    return d.getDate();
}

function _mobileGetMonthStats(ss, monthName, year, myNumbers, monthIndex) {
    var sheetName = monthName + ' ' + year;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return null;

    var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    var mIdx = (typeof monthIndex === 'number' && monthIndex >= 0) ? monthIndex : months.indexOf(monthName);
    var daysInMonth = (mIdx >= 0) ? new Date(year, mIdx + 1, 0).getDate() : 31;

    var firstRow = myNumbers.expenseFirstRow;
    var lastRow = myNumbers.expenseLastRow;
    var data = sheet.getRange(firstRow, 1, lastRow - firstRow + 1, myNumbers.expenseAmountColumn).getValues();

    var totalSpend = 0;
    var highestSpend = { description: 'None', amount: 0, category: '' };
    var categoryTotals = {};
    var entries = [];

    data.forEach(function (row) {
        var category = (row[myNumbers.expenseTypeColumn - 1] || 'Uncategorized').toString().trim();
        var description = (row[myNumbers.expenseDescrColumn - 1] || '').toString().trim();
        var dateVal = row[myNumbers.expenseDateColumn - 1];
        var amount = parseFloat(row[myNumbers.expenseAmountColumn - 1]) || 0;

        if (amount > 0) {
            totalSpend += amount;
            if (amount > highestSpend.amount) {
                highestSpend = { description: description, amount: amount, category: category };
            }
            categoryTotals[category] = (categoryTotals[category] || 0) + amount;

            var normDay = _mobileNormalizeExpenseDay(dateVal, mIdx, year, daysInMonth);
            entries.push({
                category: category,
                description: description,
                day: normDay,
                amount: amount
            });
        }
    });

    var topCategory = { name: 'None', amount: 0 };
    for (var cat in categoryTotals) {
        if (categoryTotals[cat] > topCategory.amount) {
            topCategory = { name: cat, amount: categoryTotals[cat] };
        }
    }

    return {
        totalSpend: totalSpend,
        highestSpend: highestSpend,
        topCategory: topCategory,
        categoryTotals: categoryTotals,
        entries: entries
    };
}

function _mobileGetMonthlyComparisonData(ss, currentMonthIndex, currentYear, prevMonthIndex, prevMonthYear, myNumbers, months, customAssumptions) {
    var currentMonthName = months[currentMonthIndex];
    var prevMonthName = months[prevMonthIndex];

    // Use the passed ss for the current year — also seed the cache so subsequent
    // calls to _mobileGetSpreadsheetForYear(currentYear) return it immediately.
    if (!_mobileSSCache[currentYear]) _mobileSSCache[currentYear] = ss;
    var currentSS = ss;
    var currentStats = _mobileGetMonthStats(currentSS, currentMonthName, currentYear, myNumbers, currentMonthIndex);

    var prevSS = (prevMonthYear === currentYear) ? currentSS : _mobileGetSpreadsheetForYear(prevMonthYear);
    var prevStats = prevSS ? _mobileGetMonthStats(prevSS, prevMonthName, prevMonthYear, myNumbers, prevMonthIndex) : null;

    var yearAgoYear = currentYear - 1;
    var yearAgoSS = _mobileGetSpreadsheetForYear(yearAgoYear);
    var yearAgoStats = yearAgoSS ? _mobileGetMonthStats(yearAgoSS, currentMonthName, yearAgoYear, myNumbers, currentMonthIndex) : null;

    // Current date & cutoff within the month
    var now = new Date();
    var daysInMonth = new Date(currentYear, currentMonthIndex + 1, 0).getDate();
    var currentDay = Math.min(now.getDate(), daysInMonth);

    // 1. Current month entered bills
    var currentEnteredKeys = {};
    var actualSpendToDate = currentStats ? currentStats.totalSpend : 0;

    if (currentStats && currentStats.entries) {
        currentStats.entries.forEach(function (e) {
            var desc = (e.description || '').toLowerCase().trim();
            var cat = (e.category || '').toLowerCase().trim();
            if (desc) currentEnteredKeys[desc] = true;
            if (cat) currentEnteredKeys[cat] = true;
        });
    }

    // 2. Analyze prior completed months in the current year
    var historicalBills = {};
    var pastCompletedMonthsCount = 0;

    for (var m = 0; m < currentMonthIndex; m++) {
        var pastStats = _mobileGetMonthStats(currentSS, months[m], currentYear, myNumbers, m);
        if (pastStats && pastStats.entries && pastStats.entries.length > 0) {
            pastCompletedMonthsCount++;
            pastStats.entries.forEach(function (e) {
                var desc = (e.description || '').toLowerCase().trim();
                var cat = (e.category || '').toLowerCase().trim();
                var key = desc || cat;
                if (!key) return;

                if (!historicalBills[key]) {
                    historicalBills[key] = {
                        name: e.description || e.category,
                        category: e.category,
                        days: [],
                        amounts: [],
                        monthsSeen: {}
                    };
                }
                historicalBills[key].days.push(e.day);
                historicalBills[key].amounts.push(e.amount);
                historicalBills[key].monthsSeen[m] = true;
            });
        }
    }

    // Fallback: If January (no prior completed months in current year), check prior year
    if (pastCompletedMonthsCount === 0 && yearAgoSS) {
        for (var pyM = 0; pyM < 12; pyM++) {
            var pyStats = _mobileGetMonthStats(yearAgoSS, months[pyM], yearAgoYear, myNumbers, pyM);
            if (pyStats && pyStats.entries && pyStats.entries.length > 0) {
                pastCompletedMonthsCount++;
                pyStats.entries.forEach(function (e) {
                    var desc = (e.description || '').toLowerCase().trim();
                    var cat = (e.category || '').toLowerCase().trim();
                    var key = desc || cat;
                    if (!key) return;

                    if (!historicalBills[key]) {
                        historicalBills[key] = {
                            name: e.description || e.category,
                            category: e.category,
                            days: [],
                            amounts: [],
                            monthsSeen: {}
                        };
                    }
                    historicalBills[key].days.push(e.day);
                    historicalBills[key].amounts.push(e.amount);
                    historicalBills[key].monthsSeen[pyM] = true;
                });
            }
        }
    }

    // 3. Find bills that typically appear after day D and have not been entered yet
    var projectedRemainingPosted = 0;
    var expectedRemainingBills = [];

    if (pastCompletedMonthsCount > 0) {
        for (var billKey in historicalBills) {
            var bill = historicalBills[billKey];
            var billDesc = (bill.name || '').toLowerCase().trim();
            var billCat = (bill.category || '').toLowerCase().trim();

            if (currentEnteredKeys[billKey] || (billDesc && currentEnteredKeys[billDesc]) || (billCat && currentEnteredKeys[billCat])) {
                continue;
            }

            var sumDays = 0;
            for (var dIdx = 0; dIdx < bill.days.length; dIdx++) {
                sumDays += bill.days[dIdx];
            }
            var avgDay = sumDays / bill.days.length;

            if (avgDay > currentDay) {
                var totalAmount = 0;
                for (var aIdx = 0; aIdx < bill.amounts.length; aIdx++) {
                    totalAmount += bill.amounts[aIdx];
                }
                var avgOccurAmount = totalAmount / bill.amounts.length;
                var monthsCount = Object.keys(bill.monthsSeen).length;
                var isRegular = monthsCount >= Math.ceil(pastCompletedMonthsCount / 2);
                var billProjectedAmount = isRegular ? avgOccurAmount : (totalAmount / pastCompletedMonthsCount);

                projectedRemainingPosted += billProjectedAmount;
                expectedRemainingBills.push({
                    name: bill.name,
                    category: bill.category,
                    expectedDay: Math.round(avgDay),
                    amount: billProjectedAmount
                });
            }
        }
    }

    // 4. Non-posted monthly projections from shared data (or custom assumptions)
    var groceriesMonthly = (customAssumptions && customAssumptions.groceries !== undefined)
        ? customAssumptions.groceries
        : (myNumbers.agentGroceriesMonthly || 0);
    var onlinePurchasesMonthly = (customAssumptions && customAssumptions.onlinePurchases !== undefined)
        ? customAssumptions.onlinePurchases
        : (myNumbers.agentOnlinePurchasesMonthly || 0);
    var gasolineMonthly = (customAssumptions && customAssumptions.gasoline !== undefined)
        ? customAssumptions.gasoline
        : (myNumbers.agentGasolineMonthly || 0);
    var miscMonthly = (customAssumptions && customAssumptions.misc !== undefined)
        ? customAssumptions.misc
        : (myNumbers.agentMiscMonthly || 0);
    var nonPostedMonthly = groceriesMonthly + onlinePurchasesMonthly + gasolineMonthly + miscMonthly;

    // 5. Final monthly spend forecast
    var estimatedCurrentSpend = actualSpendToDate + projectedRemainingPosted + nonPostedMonthly;

    var projectionLabel = expectedRemainingBills.length > 0
        ? 'actual + ' + expectedRemainingBills.length + ' expected bills + non-posted'
        : 'actual + non-posted assumptions';

    var vsPrevMonth = null;
    var vsYearAgo = null;

    if (currentStats && prevStats && prevStats.totalSpend > 0) {
        var prevTotalWithNonPosted = prevStats.totalSpend + nonPostedMonthly;
        var diff1 = estimatedCurrentSpend - prevTotalWithNonPosted;
        vsPrevMonth = {
            difference: diff1,
            percentChange: ((diff1 / prevTotalWithNonPosted) * 100).toFixed(1)
        };
    }

    if (currentStats && yearAgoStats && yearAgoStats.totalSpend > 0) {
        var yearAgoTotalWithNonPosted = yearAgoStats.totalSpend + nonPostedMonthly;
        var diff2 = estimatedCurrentSpend - yearAgoTotalWithNonPosted;
        vsYearAgo = {
            difference: diff2,
            percentChange: ((diff2 / yearAgoTotalWithNonPosted) * 100).toFixed(1)
        };
    }

    return {
        current: currentStats,
        currentMonthName: currentMonthName,
        previous: prevStats,
        previousMonthName: prevMonthName,
        previousMonthYear: prevMonthYear,
        yearAgo: yearAgoStats,
        yearAgoYear: yearAgoYear,
        vsPrevMonth: vsPrevMonth,
        vsYearAgo: vsYearAgo,
        projection: {
            currentDay: currentDay,
            daysInMonth: daysInMonth,
            label: projectionLabel,
            estimatedMonthlySpend: estimatedCurrentSpend,
            actualSpendToDate: actualSpendToDate,
            projectedRemainingPosted: projectedRemainingPosted,
            nonPostedMonthly: nonPostedMonthly,
            expectedRemainingBills: expectedRemainingBills
        }
    };
}

function _mobileCalculateAnnualForecast(ss, currentMonthIndex, currentYear, myNumbers, months) {
    // Use the passed ss directly — avoids an extra _getSpreadsheet() round-trip.
    // Batch-read all 12 dashboard rows in a single getValues() call.
    var dashboard = ss.getSheets()[0];
    var dashValues = dashboard
        .getRange(myNumbers.dashFirstMonthRow, myNumbers.dashAmountTotalBeforeSplitColumn, 12, 1)
        .getValues();
    var ytdPosted = 0;
    for (var i = 0; i < 12; i++) {
        var val = parseFloat(dashValues[i][0]) || 0;
        if (i <= currentMonthIndex) {
            ytdPosted += val;
        }
    }

    var prevYear = currentYear - 1;
    var prevYearSS = _mobileGetSpreadsheetForYear(prevYear);
    var prevYearTotal = 0;
    var prevYearMonthlyAvg = 0;

    if (prevYearSS) {
        // Batch-read all 12 prior-year dashboard rows in a single getValues() call.
        var prevDashboard = prevYearSS.getSheets()[0];
        var prevDashValues = prevDashboard
            .getRange(myNumbers.dashFirstMonthRow, myNumbers.dashAmountTotalBeforeSplitColumn, 12, 1)
            .getValues();
        for (var j = 0; j < 12; j++) {
            prevYearTotal += parseFloat(prevDashValues[j][0]) || 0;
        }
        prevYearMonthlyAvg = prevYearTotal / 12;
    }

    var groceriesMonthly = myNumbers.agentGroceriesMonthly || 0;
    var onlinePurchasesMonthly = myNumbers.agentOnlinePurchasesMonthly || 0;
    var gasolineMonthly = myNumbers.agentGasolineMonthly || 0;
    var miscMonthly = myNumbers.agentMiscMonthly || 0;
    var nonPostedMonthly = groceriesMonthly + onlinePurchasesMonthly + gasolineMonthly + miscMonthly;

    var remainingMonths = 11 - currentMonthIndex;
    var avgPostedThisYear = currentMonthIndex > 0 ? ytdPosted / (currentMonthIndex + 1) : 0;
    var projectedPostedPerMonth = avgPostedThisYear > 0 ? avgPostedThisYear : prevYearMonthlyAvg;
    var projectedRemaining = remainingMonths * (projectedPostedPerMonth + nonPostedMonthly);
    var annualForecast = ytdPosted + projectedRemaining;

    var yoyForecastDiff = prevYearTotal > 0 ? annualForecast - prevYearTotal : null;
    var yoyForecastPct = prevYearTotal > 0 ? ((yoyForecastDiff / prevYearTotal) * 100).toFixed(1) : null;

    return {
        ytdPosted: ytdPosted,
        monthsCompleted: currentMonthIndex + 1,
        remainingMonths: remainingMonths,
        avgMonthlyThisYear: avgPostedThisYear,
        nonPostedMonthly: nonPostedMonthly,
        nonPostedBreakdown: {
            groceries: groceriesMonthly,
            onlinePurchases: onlinePurchasesMonthly,
            gasoline: gasolineMonthly,
            misc: miscMonthly
        },
        projectedRemaining: projectedRemaining,
        annualForecast: annualForecast,
        previousYearTotal: prevYearTotal,
        vsLastYear: {
            difference: yoyForecastDiff,
            percentChange: yoyForecastPct
        }
    };
}

function _mobileDetectExpenseSpikes(currentStats, yearAgoStats, myNumbers) {
    var spikes = [];
    var aboveNormal = [];

    if (!currentStats || !yearAgoStats) {
        return { spikes: spikes, aboveNormal: aboveNormal, hasAnomalies: false };
    }

    var currentCategories = currentStats.categoryTotals || {};
    var yearAgoCategories = yearAgoStats.categoryTotals || {};

    for (var category in currentCategories) {
        var currentAmount = currentCategories[category];
        var yearAgoAmount = yearAgoCategories[category] || 0;

        if (yearAgoAmount > 0) {
            var percentChange = ((currentAmount - yearAgoAmount) / yearAgoAmount) * 100;
            if (percentChange > 50) {
                spikes.push({ category: category, currentAmount: currentAmount, yearAgoAmount: yearAgoAmount, percentChange: percentChange.toFixed(1), severity: 'HIGH' });
            } else if (percentChange > 30) {
                aboveNormal.push({ category: category, currentAmount: currentAmount, yearAgoAmount: yearAgoAmount, percentChange: percentChange.toFixed(1), severity: 'ABOVE_NORMAL' });
            }
        } else if (currentAmount > 500) {
            spikes.push({ category: category, currentAmount: currentAmount, yearAgoAmount: 0, percentChange: 'NEW', severity: 'NEW_CATEGORY' });
        }
    }

    return { spikes: spikes, aboveNormal: aboveNormal, hasAnomalies: spikes.length > 0 || aboveNormal.length > 0 };
}

function _mobileGenerateAgentAnalysis(comparisonData, forecastData, spikeAnalysis) {
    var myNumbers = new staticNumbers();
    var props = PropertiesService.getScriptProperties().getProperties();
    var apiKey = props['GEMINI_API_KEY'];

    if (!apiKey) {
        // Return diagnostic info so the UI shows exactly what property names exist
        var keys = Object.keys(props).join(', ');
        return '<strong>API key not found.</strong> Script Properties has these keys: [' + keys + ']. ' +
            'Please ensure a property named exactly <code>GEMINI_API_KEY</code> exists.';
    }

    function fmt(val) {
        return new Intl.NumberFormat('en-CA', { style: 'currency', currency: 'CAD' }).format(val || 0);
    }

    var proj = comparisonData.projection || {};
    var cur = comparisonData.current || {};

    var prompt = 'You are a household expense analysis agent. Analyze this data and provide 3-4 actionable insights.\n' +
        'Be concise, helpful, and use bold text for key numbers. Format as clean HTML list (<ul><li>...</li></ul>).\n' +
        'DO NOT use markdown code blocks like ```html. Return ONLY the HTML content.\n' +
        'Also advise of assumptions made in html (projections only, list assumptions for gasoline, misc, groceries, online purchases)\n\n' +
        'CURRENT MONTH: ' + comparisonData.currentMonthName + '\n' +
        '- Actual Spend to Date (Day ' + proj.currentDay + ' of ' + proj.daysInMonth + '): ' + fmt(proj.actualSpendToDate) + '\n' +
        '- Estimated Monthly Spend: ' + fmt(proj.estimatedMonthlySpend) + ' (' + proj.label + ')\n' +
        '- Top Category: ' + (cur.topCategory ? cur.topCategory.name : 'N/A') + ' (' + fmt(cur.topCategory ? cur.topCategory.amount : 0) + ')\n' +
        '- Highest Single Expense: ' + (cur.highestSpend ? cur.highestSpend.description : 'N/A') + ' (' + fmt(cur.highestSpend ? cur.highestSpend.amount : 0) + ')\n\n' +
        'COMPARISONS (based on estimated monthly spend):\n' +
        '- vs Previous Month (' + comparisonData.previousMonthName + '): ' + (comparisonData.vsPrevMonth ? comparisonData.vsPrevMonth.percentChange + '%' : 'N/A') + '\n' +
        '- vs Same Month Last Year: ' + (comparisonData.vsYearAgo ? comparisonData.vsYearAgo.percentChange + '%' : 'N/A') + '\n\n' +
        'ANNUAL FORECAST:\n' +
        '- YTD Posted: ' + fmt(forecastData.ytdPosted) + '\n' +
        '- Projected Annual Total: ' + fmt(forecastData.annualForecast) + '\n' +
        '- vs Last Year Annual: ' + (forecastData.vsLastYear && forecastData.vsLastYear.percentChange ? forecastData.vsLastYear.percentChange + '%' : 'N/A') + '\n\n' +
        'EXPENSE ANOMALIES:';

    if (spikeAnalysis.hasAnomalies) {
        spikeAnalysis.spikes.forEach(function (s) {
            prompt += '\n- HIGH SPIKE: ' + s.category + ' at ' + fmt(s.currentAmount) + ' (' + s.percentChange + '% vs last year)';
        });
        spikeAnalysis.aboveNormal.forEach(function (s) {
            prompt += '\n- Above Normal: ' + s.category + ' at ' + fmt(s.currentAmount) + ' (' + s.percentChange + '% vs last year)';
        });
    } else {
        prompt += '\n- No significant anomalies detected.';
    }

    prompt += '\n\nProvide insights focusing on: spending trends, areas of concern, and actionable recommendations.';

    try {
        var url = myNumbers.agentUrl + apiKey;
        var payload = {
            contents: [{ parts: [{ text: prompt }] }],
            generationConfig: { temperature: 0.7, topK: 40, topP: 0.95, maxOutputTokens: 4096 }
        };
        var options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };
        var response = UrlFetchApp.fetch(url, options);
        var json = JSON.parse(response.getContentText());

        if (response.getResponseCode() === 200 && json.candidates && json.candidates[0]) {
            var candidate = json.candidates[0];
            // Concatenate all parts (response can be split across multiple parts)
            var text = (candidate.content.parts || []).map(function (p) { return p.text || ''; }).join('').trim();
            Logger.log('Gemini finishReason: ' + candidate.finishReason + ' | length: ' + text.length);
            // Strip markdown code blocks if present
            text = text.replace(/^```html\s*/i, '').replace(/```\s*$/i, '').trim();
            return text;
        } else {
            return 'API Error (' + response.getResponseCode() + '): ' + (json.error ? json.error.message : 'Unknown response format');
        }
    } catch (e) {
        Logger.log('Gemini call failed: ' + e);
        return 'Execution Error: ' + e.toString();
    }
}
