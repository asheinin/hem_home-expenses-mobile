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

        // Spouse names
        var sp1Name = (dashboard.getRange(myNumbers.dashNamesRow, myNumbers.dashSpouse1NameColumn).getValue() || '').toString().trim();
        var sp2Name = (dashboard.getRange(myNumbers.dashNamesRow, myNumbers.dashSpouse2NameColumn).getValue() || '').toString().trim();

        // Read current month's expense sheet directly for running total
        var monthSheet = ss.getSheets()[currentMonthIdx];
        var runningTotal = 0;
        if (monthSheet) {
            var numRows = myNumbers.expenseLastRow - myNumbers.expenseFirstRow + 1;
            var values = monthSheet
                .getRange(myNumbers.expenseFirstRow, myNumbers.expenseAmountColumn, numRows, 1)
                .getValues();
            values.forEach(function (row) {
                var amt = parseFloat(row[0]);
                if (!isNaN(amt) && amt > 0) runningTotal += amt;
            });
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

        var aiAnalysis = _mobileGenerateAgentAnalysis(comparisonData, forecastData, spikeAnalysis);

        var results = {
            timestamp: now.toISOString(),
            currentMonth: currentMonthName,
            currentYear: currentYear,
            comparison: comparisonData,
            forecast: forecastData,
            spikes: spikeAnalysis,
            aiInsights: aiAnalysis
        };

        return { success: true, results: results };
    } catch (err) {
        Logger.log('runMobileExpenseAnalysis error: ' + err);
        return { success: false, message: err.toString() };
    }
}


// ============================================================================
// Internal helpers (mobile-adapted versions of agent helpers)
// ============================================================================

function _mobileGetSpreadsheetForYear(year) {
    var ss = _getSpreadsheet();
    if (ss.getName().indexOf(year.toString()) !== -1) return ss;

    // Search Drive for "Home Expenses {YYYY}" (e.g. "Home Expenses 2025")
    var targetName = 'Home Expenses ' + year;
    var files = DriveApp.getFilesByName(targetName);

    while (files.hasNext()) {
        var file = files.next();
        if (!file.isTrashed()) {
            return SpreadsheetApp.openById(file.getId());
        }
    }

    // In case there are older files named "Home payments {YYYY}"
    var oldTargetName = 'Home payments ' + year;
    var oldFiles = DriveApp.getFilesByName(oldTargetName);

    while (oldFiles.hasNext()) {
        var oldFile = oldFiles.next();
        if (!oldFile.isTrashed()) {
            return SpreadsheetApp.openById(oldFile.getId());
        }
    }

    return null;
}

function _mobileGetMonthStats(ss, monthName, year, myNumbers) {
    var sheetName = monthName + ' ' + year;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return null;

    var firstRow = myNumbers.expenseFirstRow;
    var lastRow = myNumbers.expenseLastRow;
    var data = sheet.getRange(firstRow, 1, lastRow - firstRow + 1, myNumbers.expenseAmountColumn).getValues();

    var totalSpend = 0;
    var highestSpend = { description: 'None', amount: 0, category: '' };
    var categoryTotals = {};

    data.forEach(function (row) {
        var category = (row[myNumbers.expenseTypeColumn - 1] || 'Uncategorized').toString().trim();
        var description = (row[myNumbers.expenseDescrColumn - 1] || '').toString().trim();
        var amount = parseFloat(row[myNumbers.expenseAmountColumn - 1]) || 0;

        if (amount > 0) {
            totalSpend += amount;
            if (amount > highestSpend.amount) {
                highestSpend = { description: description, amount: amount, category: category };
            }
            categoryTotals[category] = (categoryTotals[category] || 0) + amount;
        }
    });

    var topCategory = { name: 'None', amount: 0 };
    for (var cat in categoryTotals) {
        if (categoryTotals[cat] > topCategory.amount) {
            topCategory = { name: cat, amount: categoryTotals[cat] };
        }
    }

    return { totalSpend: totalSpend, highestSpend: highestSpend, topCategory: topCategory, categoryTotals: categoryTotals };
}

function _mobileGetMonthlyComparisonData(ss, currentMonthIndex, currentYear, prevMonthIndex, prevMonthYear, myNumbers, months) {
    var currentMonthName = months[currentMonthIndex];
    var prevMonthName = months[prevMonthIndex];

    var currentSS = _mobileGetSpreadsheetForYear(currentYear);
    var currentStats = currentSS ? _mobileGetMonthStats(currentSS, currentMonthName, currentYear, myNumbers) : null;

    var prevSS = (prevMonthYear === currentYear) ? currentSS : _mobileGetSpreadsheetForYear(prevMonthYear);
    var prevStats = prevSS ? _mobileGetMonthStats(prevSS, prevMonthName, prevMonthYear, myNumbers) : null;

    var yearAgoYear = currentYear - 1;
    var yearAgoSS = _mobileGetSpreadsheetForYear(yearAgoYear);
    var yearAgoStats = yearAgoSS ? _mobileGetMonthStats(yearAgoSS, currentMonthName, yearAgoYear, myNumbers) : null;

    // Projection multiplier
    var now = new Date();
    var currentDay = now.getDate();
    var daysInMonth = new Date(currentYear, currentMonthIndex + 1, 0).getDate();
    var oneThird = daysInMonth / 3;
    var oneHalf = daysInMonth / 2;

    var projectionMultiplier = 1;
    var projectionLabel = 'actual (past mid-month)';
    if (currentDay <= oneThird) {
        projectionMultiplier = 2;
        projectionLabel = '2× estimate (first 1/3 of month)';
    } else if (currentDay <= oneHalf) {
        projectionMultiplier = 1.2;
        projectionLabel = '1.2× estimate (between 1/3 and 1/2 of month)';
    }

    var estimatedCurrentSpend = currentStats ? currentStats.totalSpend * projectionMultiplier : 0;

    var vsPrevMonth = null;
    var vsYearAgo = null;

    if (currentStats && prevStats && prevStats.totalSpend > 0) {
        var diff1 = estimatedCurrentSpend - prevStats.totalSpend;
        vsPrevMonth = {
            difference: diff1,
            percentChange: ((diff1 / prevStats.totalSpend) * 100).toFixed(1)
        };
    }

    if (currentStats && yearAgoStats && yearAgoStats.totalSpend > 0) {
        var diff2 = estimatedCurrentSpend - yearAgoStats.totalSpend;
        vsYearAgo = {
            difference: diff2,
            percentChange: ((diff2 / yearAgoStats.totalSpend) * 100).toFixed(1)
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
            multiplier: projectionMultiplier,
            label: projectionLabel,
            estimatedMonthlySpend: estimatedCurrentSpend,
            actualSpendToDate: currentStats ? currentStats.totalSpend : 0
        }
    };
}

function _mobileCalculateAnnualForecast(ss, currentMonthIndex, currentYear, myNumbers, months) {
    var dashboard = _getSpreadsheet().getSheets()[0];
    var ytdPosted = 0;

    for (var i = 0; i < 12; i++) {
        var row = myNumbers.dashFirstMonthRow + i;
        var val = parseFloat(dashboard.getRange(row, myNumbers.dashAmountTotalBeforeSplitColumn).getValue()) || 0;
        if (i <= currentMonthIndex) {
            ytdPosted += val;
        }
    }

    var prevYear = currentYear - 1;
    var prevYearSS = _mobileGetSpreadsheetForYear(prevYear);
    var prevYearTotal = 0;
    var prevYearMonthlyAvg = 0;

    if (prevYearSS) {
        var prevDashboard = prevYearSS.getSheets()[0];
        for (var j = 0; j < 12; j++) {
            var pRow = myNumbers.dashFirstMonthRow + j;
            var pVal = parseFloat(prevDashboard.getRange(pRow, myNumbers.dashAmountTotalBeforeSplitColumn).getValue()) || 0;
            prevYearTotal += pVal;
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
