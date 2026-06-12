class staticNumbers {
  constructor() {
    this.thresholdLimitForClosingMonth = 1;

    this.expenseCarryOverRow = 2;
    this.expenseSpToSpRow = 2;
    this.expenseFirstRow = 3;
    this.expenseLastRow = 50;
    this.expenseTotalAmountRowBeforeSplit = 52;
    this.expenseTotalAmountRow = 53;
    this.expenseTotalPaidSp2Row = 56;
    this.expenseTotalPaidSp1Row = 60;
    this.expenseSp2MonthlyBalanceRow = 63;
    this.expenseSp1MonthlyBalanceRow = 64;
    this.expenseSp2InitBalancePaidRow = 66;
    this.expenseSp2InitBalanceLeftRow = 67;
    this.expenseSp1InitBalancePaidRow = 69;
    this.expenseSp1InitBalanceLeftRow = 70;

    this.expenseTypeColumn = 1;
    this.expenseDescrColumn = 2;
    this.expenseDateColumn = 3;
    this.expenseAmountColumn = 4;
    this.expenseCarryOverSp2OwesColumn = 5;
    this.expenseCarryOverSp1OwesColumn = 6;
    this.expenceSplit2Column = 5;
    this.expenceSplit1Column = 6;
    this.expenceSplitColumn = 7;
    this.expenseFirstPayColumn = 8;
    this.expenseSecondPayColumn = 9;
    this.expenseInitialBalanceCol = 2;
    this.expencePeriodColumn = 10;
    this.expensePaidColumn = 11;
    this.expensePAPColumn = 12;
    this.expenseSp2ToSp1Column = 13;
    this.expenseSp1ToSp2Column = 14;

    // Dashboard rows

    this.dashAddressRow = 1;

    this.dashNamesRow = 2;
    this.dashEmailsRow = 3;
    this.dashSplitRow = 4;
    this.dashBalancesRow = 5;
    this.dashTitleRow = 6;
    this.dashFirstMonthRow = 7;

    this.dashSpouse1NameColumn = 2;
    this.dashSpouse2NameColumn = 3;
    this.dashAddressColumn = 2;
    this.dashSp1SplitColumn = 2;
    this.dashSp2SplitColumn = 3;

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

    // Dashboard balance cell color constants

    this.dashBalanceNegativeBgColor = "red";
    this.dashBalancePositiveBgColor = "green";
    this.dashBalanceNeutralBgColor = "green";

    // Summary rows/columns

    this.summaryHeaderRow = 1;
    this.summarySumRow = 2;

    this.summaryAmountColumn = 3;
    this.summaryAnalyticsYearColumn = 1;
    this.summaryAnalyticsDataStartColumn = 4;
    this.summaryChartsStartColumn = 17;
    this.summaryAnalyticsMonthColumn = 1;
    this.summaryAnalyticsTotalColumns = 60;

    // Expense Analysis Agent - Non-posted monthly projections
    this.agentGroceriesMonthly = 800;
    this.agentOnlinePurchasesMonthly = 800;
    this.agentGasolineMonthly = 0;
    this.agentMiscMonthly = 1000;

    // Expense Analysis Agent - API Key
    this.agentApiKey = 'GEMINI_API_KEY';
    this.agentModel = 'gemini-2.5-flash';
    this.agentUrl = 'https://generativelanguage.googleapis.com/v1beta/models/' + this.agentModel + ':generateContent?key=';

    // Privacy Policy, Support and App Info
    this.privacyPolicyUrl = 'https://github.com/asheinin/hem_home-expenses-mobile/blob/main/PrivacyPolicy.html';
    this.supportEmail = 'galaxsolutions@gmail.com';
    this.appVersion = '0.0.1';
    this.appName = 'Home Expenses';
    this.appDeveloper = 'Galaxsolutions';

    // Global settings
    this.mailerName = 'HomePayments';
    this.sourceFilename = 'Home payments';
  }
}
