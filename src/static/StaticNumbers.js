/**
 * Static column/row constants for Home Expenses spreadsheet.
 * Kept in sync with HomeExpenses/src/static/main.js
 */
class staticNumbers {
    constructor() {
        this.expenseCarryOverRow = 2;
        this.expenseFirstRow = 3;
        this.expenseLastRow = 50;

        this.expenseTypeColumn = 1;
        this.expenseDescrColumn = 2;
        this.expenseDateColumn = 3;
        this.expenseAmountColumn = 4;
        this.expenceSplit2Column = 5;
        this.expenceSplit1Column = 6;
        this.expenceSplitColumn = 7;
        this.expenseFirstPayColumn = 8;
        this.expenseSecondPayColumn = 9;
        this.expencePeriodColumn = 10;
        this.expensePaidColumn = 11;
        this.expensePAPColumn = 12;

        // Dashboard rows
        this.dashNamesRow = 2;
        this.dashEmailsRow = 3;
        this.dashFirstMonthRow = 7;

        // Dashboard name columns
        this.dashSpouse1NameColumn = 2;
        this.dashSpouse2NameColumn = 3;

        // Dashboard data columns
        this.dashMonthNameColumn = 1;
        this.dashSp1BalanceUsedColumn = 2;
        this.dashSp2BalanceUsedColumn = 3;
        this.dashAmountTotalBeforeSplitColumn = 4;
        this.dashAmountTotalColumn = 5;
        this.dashSp2PartColumn = 6;
        this.dashSp1PartColumn = 7;
        this.dashSp2PaidColumn = 8;
        this.dashSp1PaidColumn = 9;
        this.dashSp2ToSp1Column = 10;
        this.dashSp1ToSp2Column = 11;
        this.dashSp2BalanceColumn = 12;
        this.dashSp1BalanceColumn = 13;
        this.dashColumns = 13;

        // Expense Analysis Agent — non-posted monthly projection defaults
        this.agentGroceriesMonthly = 800;
        this.agentOnlinePurchasesMonthly = 800;
        this.agentGasolineMonthly = 0;
        this.agentMiscMonthly = 1000;
    }
}
