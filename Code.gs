function onOpen() {
  try {
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Expense Manager");
    menu.addItem("Setup", "setup");
    menu.addItem("Create new expense sheet", "createExpenseSheet");

    let subMenu = ui.createMenu("Category-wise report");
    subMenu.addItem("Current sheet", "categoryWiseReportCurrentSheet");
    subMenu.addItem("All sheets", "categoryWiseReportAllSheets");
    menu.addSubMenu(subMenu);

    subMenu = ui.createMenu("Payer-wise report");
    subMenu.addItem("Current sheet", "payerWiseReportCurrentSheet");
    subMenu.addItem("All sheets", "payerWiseReportAllSheets");
    menu.addSubMenu(subMenu);

    subMenu = ui.createMenu("Monthly report");
    subMenu.addItem("Current sheet", "monthlyReportCurrentSheet");
    subMenu.addItem("All sheets", "monthlyReportAllSheets");
    menu.addSubMenu(subMenu);

    subMenu = ui.createMenu("Yearly report");
    subMenu.addItem("Current sheet", "yearlyReportCurrentSheet");
    subMenu.addItem("All sheets", "yearlyReportAllSheets");
    menu.addSubMenu(subMenu);

    menu.addToUi();
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

/********** Start: UI Callback Functions **********/
function setup() {
  try {
    let expenseMgrApp = new ExpenseManagerApp();
    expenseMgrApp.setup();
    Browser.msgBox("Success", "Setup completed", Browser.Buttons.OK);
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

function categoryWiseReportCurrentSheet() {
  try {
    let expenseMgrApp = new ExpenseManagerApp();
    expenseMgrApp.categoryWiseReportCurrentSheet();
    Browser.msgBox("Success", "Summary generation completed", Browser.Buttons.OK);
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

function categoryWiseReportAllSheets() {
  try {
    let expenseMgrApp = new ExpenseManagerApp();
    expenseMgrApp.categoryWiseReportAllSheets();
    Browser.msgBox("Success", "Summary generation completed", Browser.Buttons.OK);
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

function monthlyReportCurrentSheet() {
  try {
    let expenseMgrApp = new ExpenseManagerApp();
    expenseMgrApp.monthlyReportCurrentSheet();
    Browser.msgBox("Success", "Monthly report generation completed", Browser.Buttons.OK);
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

function monthlyReportAllSheets() {
  try {
    let expenseMgrApp = new ExpenseManagerApp();
    expenseMgrApp.monthlyReportAllSheets();
    Browser.msgBox("Success", "Monthly report generation completed", Browser.Buttons.OK);
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

function yearlyReportCurrentSheet() {
  try {
    let expenseMgrApp = new ExpenseManagerApp();
    expenseMgrApp.yearlyReportCurrentSheet();
    Browser.msgBox("Success", "Yearly report generation completed", Browser.Buttons.OK);
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

function yearlyReportAllSheets() {
  try {
    let expenseMgrApp = new ExpenseManagerApp();
    expenseMgrApp.yearlyReportAllSheets();
    Browser.msgBox("Success", "Yearly report generation completed", Browser.Buttons.OK);
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

function payerWiseReportCurrentSheet() {
  try {
    let expenseMgrApp = new ExpenseManagerApp();
    expenseMgrApp.payerWiseReportCurrentSheet();
    Browser.msgBox("Success", "Payer spending report generated", Browser.Buttons.OK);
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

function payerWiseReportAllSheets() {
  try {
    let expenseMgrApp = new ExpenseManagerApp();
    expenseMgrApp.payerWiseReportAllSheets();
    Browser.msgBox("Success", "Payer spending report generated", Browser.Buttons.OK);
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

function createExpenseSheet() {
  try {
    let expenseMgrApp = new ExpenseManagerApp();
    const newSheet = expenseMgrApp.createExpenseSheet();
    Browser.msgBox("Success", `Created new expense sheet '${newSheet.getSheetName()}'`, Browser.Buttons.OK);
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

/********** End: UI Callback Functions **********/

class ExpenseManagerApp {
  /********** Start: Public functions **********/
  constructor() {
    this._init();
  }

  setup() {
    this._setupImpl();
  }

  categoryWiseReportCurrentSheet() {
    this._generateCategoryWiseReport([this._getCurrentSheetName()]);
  }

  categoryWiseReportAllSheets() {
    this._generateCategoryWiseReport(this._expenseSheetNames);
  }

  monthlyReportCurrentSheet() {
    this._monthlyReport([this._getCurrentSheetName()]);
  }
  
  monthlyReportAllSheets() {
    this._monthlyReport(this._expenseSheetNames);
  }

  yearlyReportCurrentSheet() {
    this._yearlyReport([this._getCurrentSheetName()]);
  }

  yearlyReportAllSheets() {
    this._yearlyReport(this._expenseSheetNames);
  }

  payerWiseReportCurrentSheet() {
    this._payerReport([this._getCurrentSheetName()]);
  }

  payerWiseReportAllSheets() {
    this._payerReport(this._expenseSheetNames);
  }

  createExpenseSheet() {
    return this._createExpenseSheetImpl();
  }
  
  /********** End: Public functions **********/

  /********** Start: Private functions **********/
  _init() {
    this._spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    this._selectionsSheetName = "Selections";
    this._categoryWiseReportSheetName = "Category Spending Report";
    this._monthlyReportSheetName = "Monthly Breakdown";
    this._payerReportSheetName = "Payer Spending Report";
    this._yearlyReportSheetName = "Yearly Breakdown";
    this._NonExpenseSheetNames = new Set([this._selectionsSheetName, this._categoryWiseReportSheetName, 
                                          this._monthlyReportSheetName, this._payerReportSheetName,
                                          this._yearlyReportSheetName]);

    this._dateColName = "Date";
    this._amountColName = "Amount (INR)";
    this._expCatColName = "Expense Category";
    this._expDescColName = "Expense Description";
    this._expGrpColName = "Expense Group";
    this._payerColName = "Payer";
    this._payModeColName = "Payment Mode";

    this._dateFormat = "dd-MMM-yyyy";
    this._monthYearFormat = "MMM yyyy";
    this._dailyExpSheetName = "Daily Expenses";
    this._refundCatName = "Refund";
    this._fontFamily = "Consolas";
    this._invalidDate = "Invalid Date";

    // Use array for below contains to detect and error out duplicates
    this._expenseSheetNames = [];
    this._expenseCategories = [];
    this._expenseGroups = [];
    this._payers = [];
    this._paymentModes = [];

    // Use Set for 'Selections' for faster lookup
    this._expenseCategoriesSet = new Set();
    this._expenseGroupsSet = new Set();
    this._payersSet = new Set();
    this._paymentModesSet = new Set();

    this._validationErrors = [];
    this._categoryReportData = new Map();
    this._totalSum = 0;
    this._monthlyData = new Map();
    this._yearlyData = new Map();
    this._payerData = new Map();

    this._computeAllExpenseSheetNames();
    this._fetchAllSelections();

    this._categoryReportHeaderRow = [this._expGrpColName, this._expCatColName, this._amountColName, "% Within Group", "% of Overall Total"];
    this._monthlyReportHeaderRow = ["Month", ...this._expenseCategories, "Monthly Total (INR)"];
    this._yearlyReportHeaderRow = ["Year", ...this._expenseCategories, "Yearly Total (INR)"];
    this._payerSpendReportHeaderRow = ["Payer", ...this._expenseGroups, "Payer Total (INR)"];
  }

  _setupImpl() {
    this._applyStandardFormatting();
    this._validateSelections();
    this._standardizeDateFormat();
    this._sortByDate();
    this._applyNumericValidation();
    this._setupDropdowns();
  }

  _createExpenseSheetImpl() {
    const date = new Date();
    const timeZone = this._spreadSheet.getSpreadsheetTimeZone();
    const baseSheetName = Utilities.formatDate(date, timeZone, this._monthYearFormat); // e.g. "Jul 2025"
    let newSheetName = baseSheetName;
    let counter = 0;
    while(this._spreadSheet.getSheetByName(newSheetName) !== null) {
      ++counter;
      newSheetName = `${baseSheetName} (${counter})`; // e.g. "Jul 2025 (1)"
    }
    let newSheet = this._spreadSheet.insertSheet(newSheetName);
    const headerRow = [this._dateColName, this._amountColName, this._expCatColName, 
                       this._expDescColName, this._expGrpColName, this._payerColName, this._payModeColName];

    let range = newSheet.getRange(1, 1, 1, headerRow.length);
    range.setValues([headerRow]);
    range.setFontWeight("bold");

    this._applyStandardFormattingToSheet(newSheet);
    this._applyDatePickerToFirstColumnToSheet(newSheet);
    this._applyNumericValidationToSheet(newSheet);
    this._setupDrowdownsToSheet(newSheet);
    Logger.log(`Created new expense sheet ${newSheet.getSheetName()}`);

    return newSheet;
  }

  _generateCategoryWiseReport(sheetNames) {
    this._populateSummaryData(sheetNames);
    this._generateCategoryWiseReportSheet();
  }

  _monthlyReport(sheetNames) {
    this._populateMonthlyReportData(sheetNames);
    this._generateMonthlyReportSheet();
  }

  _yearlyReport(sheetNames) {
    this._populateYearlyReportData(sheetNames);
    this._generateYearlyReportSheet();
  }

  _payerReport(sheetNames) {
    this._populatePayerData(sheetNames);
    this._generatePayerSpendingReport();
  }

  _getCurrentSheetName() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    return sheetName;
  }

  _applyStandardFormattingToSheet(sheet) {
    Logger.log(`Apply formatting to sheet ${sheet.getName()}`);
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    const lastCol = sheet.getLastColumn() > 0 ? sheet.getLastColumn() : 1;
    let range = sheet.getRange(1, 1, maxRows, maxCols);
    range.setFontFamily(this._fontFamily);
    for(let iCol = 1; iCol <= lastCol; ++iCol) {
      sheet.autoResizeColumn(iCol);
      // Further increase the size by 20%
      sheet.setColumnWidth(iCol, sheet.getColumnWidth(iCol) * 1.2);
    }
    sheet.setFrozenRows(1);
  }

  _applyStandardFormatting() {
    for(let sheet of this._spreadSheet.getSheets()) {
      this._applyStandardFormattingToSheet(sheet);
    }
    Logger.log("Finished applying the format");
  }

  _applyDatePickerToFirstColumnToSheet(sheet) {
    const colNum = 1; // date column number
    const rowNum = 2;
    const numRows = sheet.getMaxRows()-1; // skip the first row
    const numCols = 1;
    const dateRange = sheet.getRange(rowNum, colNum, numRows, numCols);
    dateRange.setNumberFormat(this._dateFormat);
    const dateRule = SpreadsheetApp.newDataValidation()
                                   .requireDate()
                                   .setAllowInvalid(false)
                                   .setHelpText("Select a date")
                                   .build();
    dateRange.setDataValidation(dateRule);
  }

  _applyNumericValidationToSheet(sheet) {
      const values = sheet.getDataRange().getValues();
      const headerRow = sheet.getDataRange().getValues()[0];
      const amountIdx = this._getColumnIndexBasedOnColumnName(headerRow, this._amountColName);
      const amountColNum = amountIdx + 1;
      const startRowNum = 2;
      const numRows = sheet.getMaxRows() - 1; // skip the header row
      const numCols = 1;

      const amtRange = sheet.getRange(startRowNum, amountColNum, numRows, numCols);
      const numberRule = SpreadsheetApp.newDataValidation()
                                       .requireNumberGreaterThan(0)
                                       .setAllowInvalid(false)
                                       .setHelpText("Enter a number greater than 0")
                                       .build();
      amtRange.setDataValidation(numberRule);
  }

  _applyNumericValidation() {
    for (const sheetName of this._expenseSheetNames) {
      Logger.log(`Applying number validation to Expense Amount in sheet ${sheetName}`);
      let sheet = this._spreadSheet.getSheetByName(sheetName);
      this._applyNumericValidationToSheet(sheet);
    }
  }

  _standardizeDateFormat() {
    let invalidDateCount = 0;
    for (const sheetName of this._expenseSheetNames) {
      Logger.log(`Standardizing date format in sheet ${sheetName}`);
      let sheet = this._spreadSheet.getSheetByName(sheetName);
      this._applyDatePickerToFirstColumnToSheet(sheet);
      const headerRow = sheet.getDataRange().getValues()[0];
      const dateIdx = this._getColumnIndexBasedOnColumnName(headerRow, this._dateColName);
      const dateColNum = dateIdx + 1;
      const dateRowNum = 2;

      const lastRow = sheet.getLastRow();
      if(lastRow <= 1) {
        // Handles sheets that are empty or contain only a header row.
        continue;
      }
      let dataRange = sheet.getRange(dateRowNum, dateColNum, lastRow - 1, 1);
      const dateValues = dataRange.getValues();

      const newDateValues = [];
      let invalidDateRowNums = [];

      for(let i = 0; i < dateValues.length; ++i) {
        const originalValue = dateValues[i][0];
        const date = new Date(originalValue);
        if(date.toString() !== this._invalidDate) {
          const timezone = this._spreadSheet.getSpreadsheetTimeZone();
          const formattedDate = Utilities.formatDate(date, timezone, this._dateFormat);
          newDateValues.push([formattedDate]);
        } else {
          Logger.log(`Row ${i + 2}: Invalid date format`);
          ++invalidDateCount;
          newDateValues.push([originalValue]);
          invalidDateRowNums.push(i + 2);
        }
      }

      dataRange.setValues(newDateValues);
      dataRange.setNumberFormat(this._dateFormat);
      dataRange.setBackground('null');
      for (const rowNum of invalidDateRowNums) {
        sheet.getRange(rowNum, dateColNum).setBackground('red');
      }
    }
    Logger.log("Date standardization complete");

    if(invalidDateCount > 0) {
      const errMsg = `Found ${invalidDateCount} cells with invalid date. \n\nCheck the highlighted cells for details`;
      let ui = SpreadsheetApp.getUi();
      ui.alert("Date Validation Error", errMsg, ui.ButtonSet.OK);
    }
  }

  _createSelectionsSheet() {
    let selectionsSheet = this._spreadSheet.getSheetByName(this._selectionsSheetName);
    selectionsSheet = this._spreadSheet.insertSheet(this._selectionsSheetName);
    let headerRow = [this._expCatColName, this._expGrpColName, this._payerColName, this._payModeColName];
    let range = selectionsSheet.getRange(1, 1, 1, headerRow.length);
    range.setValues([headerRow]);
    range.setFontWeight("bold");

    const sampleExpCategories = [
                                  "Accommodation",
                                  "Electronics",
                                  "Entertainment",
                                  "Fees",
                                  "Food",
                                  "Gift",
                                  "Groceries",
                                  "Healthcare",
                                  "Household",
                                  "Insurance",
                                  "Refund",
                                  "Rent",
                                  "Shopping",
                                  "Tax",
                                  "Transportation",
                                  "Utilities",
                                  "Vehicle"
                                ];
    const sampleExpGrp = [
                            "Daily Expenses",
                            "Dubai Vacation",
                            "Rare Expenses",
                            "USA Vacation"
                          ];
    const samplePayers = [
                            "Ethan Hunt",
                            "Jack Reacher",
                            "Jack Ryan",
                            "James Bond",
                            "Jason Bourne"
                          ];
    const samplePayModes = [
                              "Cash",
                              "Online"
                            ];
    const startRowNum = 2;
    const expCatColNum = this._getColumnIndexBasedOnColumnName(headerRow, this._expCatColName) + 1;
    const expGrpColNum = this._getColumnIndexBasedOnColumnName(headerRow, this._expGrpColName) + 1;
    const payerColNum = this._getColumnIndexBasedOnColumnName(headerRow, this._payerColName) + 1;
    const payModeColNum = this._getColumnIndexBasedOnColumnName(headerRow, this._payModeColName) + 1;

    range = selectionsSheet.getRange(startRowNum, expCatColNum, sampleExpCategories.length, 1);
    // Transform your 1D array into a 2D array required by setValues()
    // Each item needs to be wrapped in its own array: ['Apple'] becomes [['Apple']]
    range.setValues(sampleExpCategories.map(value => [value]));

    range = selectionsSheet.getRange(startRowNum, expGrpColNum, sampleExpGrp.length, 1);
    range.setValues(sampleExpGrp.map(value => [value]));

    range = selectionsSheet.getRange(startRowNum, payerColNum, samplePayers.length, 1);
    range.setValues(samplePayers.map(value => [value]));

    range = selectionsSheet.getRange(startRowNum, payModeColNum, samplePayModes.length, 1);
    range.setValues(samplePayModes.map(value => [value]));

    this._applyStandardFormattingToSheet(selectionsSheet);
    Logger.log("Created new Selections sheet");
    return selectionsSheet;
  }

  _populatePayerData(sheetNames) {
    for(const sheetName of sheetNames) {
      const sheet = this._spreadSheet.getSheetByName(sheetName);
      Logger.log(`Analyzing ${sheetName} for computing Payer Spending Report`);
      const values = sheet.getDataRange().getValues();
      const headerRow = values[0];
      const amountIdx = this._getColumnIndexBasedOnColumnName(headerRow, this._amountColName);
      const expGrpIdx = this._getColumnIndexBasedOnColumnName(headerRow, this._expGrpColName);
      const payerIdx = this._getColumnIndexBasedOnColumnName(headerRow, this._payerColName);

      for(let i = 1; i < values.length; ++i) {
        // start for index 1 as 0th index contains the header
        const currRow = values[i];
        const expGrpName = currRow[expGrpIdx].toString().trim();
        const amount = parseFloat(currRow[amountIdx]);
        const payerName = currRow[payerIdx].toString().trim();

        let payerGrpData = this._payerData.get(payerName);
        if(!payerGrpData) {
          payerGrpData = {"PayerTotal": 0, "GroupSum": new Map()};
          this._payerData.set(payerName, payerGrpData);
        }
        const currSum = payerGrpData.GroupSum.get(expGrpName) || 0;
        payerGrpData.PayerTotal += amount;
        payerGrpData.GroupSum.set(expGrpName, currSum + amount);
      }
    }
  }

  _generatePayerSpendingReport() {
    let output = [];
    output.push(this._payerSpendReportHeaderRow);

    for(const [payerName, payerInfo] of this._payerData) {
      let payerRow = [];
      payerRow.push(payerName);
      for(const expGrpName of this._expenseGroups) {
        const grpSum = payerInfo.GroupSum.get(expGrpName) || 0;
        payerRow.push(grpSum);
      }
      payerRow.push(payerInfo.PayerTotal);
      output.push(payerRow);
    }
    let payerReportSheet = this._spreadSheet.getSheetByName(this._payerReportSheetName);
    if(payerReportSheet) {
      payerReportSheet.clearContents();
    } else {
      payerReportSheet = this._spreadSheet.insertSheet(this._payerReportSheetName);
    }

    let range = payerReportSheet.getRange(1, 1, output.length, output[0].length);
    range.setValues(output);

    // Bold header row
    payerReportSheet.getRange(1, 1, 1, output[0].length).setFontWeight("bold");
    this._applyStandardFormattingToSheet(payerReportSheet);

    Logger.log(`Payer report sheet generated in ${this._payerReportSheetName} sheet`);
  }

  _populateMonthlyReportData(sheetNames) {
    const keyGenerator = (dateObj) => this._getMonthYearString(dateObj);
    this._populateTimeBasedReportData(sheetNames, keyGenerator, this._monthlyData);
  }

  _populateYearlyReportData(sheetNames) {
    const keyGenerator = (dateObj) => dateObj.getFullYear().toString();
    this._populateTimeBasedReportData(sheetNames, keyGenerator, this._yearlyData);
  }

  _populateTimeBasedReportData(sheetNames, keyGenerator, reportData) {
    for(const sheetName of sheetNames) {
      if(this._isNonExpenseSheet(sheetName)) {
        throw new Error(`Cannot generate monthly report for sheet '${sheetName}'`);
      }
      let sheet = this._spreadSheet.getSheetByName(sheetName);
      const values = sheet.getDataRange().getValues();
      const headerRow = sheet.getDataRange().getValues()[0];
      const amountIdx = this._getColumnIndexBasedOnColumnName(headerRow, this._amountColName);
      const expCatIdx = this._getColumnIndexBasedOnColumnName(headerRow, this._expCatColName);
      const dateIdx = this._getColumnIndexBasedOnColumnName(headerRow, this._dateColName);

      for(let i = 1; i < values.length; ++i) {
        // start for index 1 as 0th index contains the header
        const currRow = values[i];

        const expCatName = currRow[expCatIdx].toString().trim();
        const amount = parseFloat(currRow[amountIdx]);
        const dateStr = currRow[dateIdx].toString().trim();
        const dateObj = new Date(dateStr);
        const keyStr = keyGenerator(dateObj);

        let reportSummary = reportData.get(keyStr);
        if(!reportSummary) {
          reportSummary = {"Total": 0.0, "CategoricalSum": new Map()};
          reportData.set(keyStr, reportSummary);
        }
        if(expCatName !== this._refundCatName) {
          reportSummary.Total += amount;
        }
        const currSum = reportSummary.CategoricalSum.get(expCatName) || 0;
        reportSummary.CategoricalSum.set(expCatName, currSum + amount);
      }
    }
  }

  _generateYearlyReportSheet() {
    this._generateTimeBasedReport(this._yearlyReportHeaderRow, this._yearlyData, this._yearlyReportSheetName);
    Logger.log(`Yearly report generated in ${this._yearlyReportSheetName} sheet`);
  }

  _generateMonthlyReportSheet() {
    this._generateTimeBasedReport(this._monthlyReportHeaderRow, this._monthlyData, this._monthlyReportSheetName);
    Logger.log(`Monthly report generated in ${this._monthlyReportSheetName} sheet`);
  }

  _generateTimeBasedReport(headerRow, reportData, reportSheetName) {
    let output = [];
    output.push(headerRow);

    for(const [keyStr, reportInfo] of reportData) {
      let reportRow = [];
      reportRow.push(keyStr);
      for(const categoryName of this._expenseCategories) {
        const categorySum = reportInfo.CategoricalSum.get(categoryName) || 0;
        reportRow.push(categorySum);
      }
      reportRow.push(reportInfo.Total);
      output.push(reportRow);
    }
    let reportSheet = this._spreadSheet.getSheetByName(reportSheetName);
    if(reportSheet) {
      reportSheet.clearContents();
    } else {
      reportSheet = this._spreadSheet.insertSheet(reportSheetName);
    }

    let range = reportSheet.getRange(1, 1, output.length, output[0].length);
    range.setValues(output);

    // Bold header row
    reportSheet.getRange(1, 1, 1, output[0].length).setFontWeight("bold");
    this._applyStandardFormattingToSheet(reportSheet);
  }
  
  _getMonthYearString(date) {
    const monthNames = [
      "Jan", "Feb", "Mar", "Apr", "May", "Jun",
      "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
    ];
    const month = monthNames[date.getMonth()];
    const year = date.getFullYear();
    return `${month}-${year}`;
  }

  _validateSelections() {
    this._validationErrors = [];
    for(const sheetName of this._expenseSheetNames) {
      this._validateSelectionsInSheet(sheetName);
    }

    if(this._validationErrors.length !== 0) {
      const errMsg = this._validationErrors.join("\n") + "\n\nCheck the highlighted cells for details";
      let ui = SpreadsheetApp.getUi();
      ui.alert("Validation Error", errMsg, ui.ButtonSet.OK);
    }
  }

  _setupDrowdownsToSheet(sheet) {
    Logger.log(`Adding data validation rule for ${sheet.getSheetName()}`);
    const [ruleExpCat, ruleExpGrp, rulePayer, rulePayMode] = this._getDataValidationRules();
    const values = sheet.getDataRange().getValues();
    const headerRow = values[0];
    const [expenseCatColIdx, expenseGrpColIdx, payerColIdx, payModeColIdx]  = this._getColumnIndices(headerRow);
    
    const rowNum = 2;
    const [expCatColNum, expGrpColNum, payerColNum, payModeColNum] = 
    [expenseCatColIdx + 1, expenseGrpColIdx + 1, payerColIdx + 1, payModeColIdx + 1];
    const numRows = sheet.getMaxRows() - 1;
    const numCols = 1;
    
    const rangeExpCat = sheet.getRange(rowNum, expCatColNum, numRows, numCols);
    rangeExpCat.setDataValidation(ruleExpCat);

    const rangeExpGrp = sheet.getRange(rowNum, expGrpColNum, numRows, numCols);
    rangeExpGrp.setDataValidation(ruleExpGrp);

    const rangePayer = sheet.getRange(rowNum, payerColNum, numRows, numCols);
    rangePayer.setDataValidation(rulePayer);

    const rangePayMode = sheet.getRange(rowNum, payModeColNum, numRows, numCols);
    rangePayMode.setDataValidation(rulePayMode); 
  }

  _setupDropdowns() {
    for(const sheetName of this._expenseSheetNames) {
      let sheet = this._spreadSheet.getSheetByName(sheetName);
      this._setupDrowdownsToSheet(sheet);
    }
  }

  _generateCategoryWiseReportSheet() {
    let output = [];
    output.push(this._categoryReportHeaderRow);
    for(const [groupName, groupInfo] of this._categoryReportData) {
      const grpTotal = groupInfo.GroupTotal;
      const grpPercent = grpTotal / this._totalSum * 100;
      output.push([groupName, "", grpTotal, "", grpPercent.toFixed(2)]);

      for(const [categoryName, categorySum] of groupInfo.CategorialSum) {
        const categoryPercentWithinGrp = categorySum / grpTotal * 100;
        const categoryPercentOverall = categorySum / this._totalSum * 100;
        output.push(["", categoryName, categorySum, categoryPercentWithinGrp.toFixed(2), categoryPercentOverall.toFixed(2)]);
      }
    }
    output.push(["Grand Total", "", this._totalSum, "", 100]);
    let summarySheet = this._spreadSheet.getSheetByName(this._categoryWiseReportSheetName);
    if(summarySheet) {
      summarySheet.clearContents();
    } else {
      summarySheet = this._spreadSheet.insertSheet(this._categoryWiseReportSheetName);
    }

    let range = summarySheet.getRange(1, 1, output.length, output[0].length);
    range.setValues(output);

    summarySheet.getRange(1, 1, 1, output[0].length).setFontWeight("bold");
    this._applyStandardFormattingToSheet(summarySheet);
    Logger.log(`Expense Summary Generated in ${this._categoryWiseReportSheetName} sheet`);
  }

  _sortByDate() {
    for(const sheetName of this._expenseSheetNames) {
      Logger.log(`Sorting by date in sheet ${sheetName}`);
      const expenseSheet = this._spreadSheet.getSheetByName(sheetName);
      const lastRow = expenseSheet.getLastRow();
      if(lastRow <= 1) {
        continue;
      }
      let range = expenseSheet.getRange(2, 1, lastRow - 1, expenseSheet.getLastColumn());
      range.sort({column: 1, ascending: true});
    }
  }

  _populateSummaryDataFromSheet(sheet) {
    const values = sheet.getDataRange().getValues();
    const headerRow = values[0];
    const amountIdx = this._getColumnIndexBasedOnColumnName(headerRow, this._amountColName);
    const expCatIdx = this._getColumnIndexBasedOnColumnName(headerRow, this._expCatColName);
    const expGrpIdx = this._getColumnIndexBasedOnColumnName(headerRow, this._expGrpColName);

    for(let i = 1; i < values.length; ++i) {
      // start for index 1 as 0th index contains the header
      const currRow = values[i];

      const expGrpName = currRow[expGrpIdx].toString().trim();
      const expCatName = currRow[expCatIdx].toString().trim();
      const amount = parseFloat(currRow[amountIdx]);

      let groupSummary = this._categoryReportData.get(expGrpName);
      if(!groupSummary) {
        groupSummary = {"GroupTotal": 0.0, "CategorialSum": new Map()};
        this._categoryReportData.set(expGrpName, groupSummary);
      }
      if(expCatName !== this._refundCatName) {
        groupSummary.GroupTotal += amount;
        this._totalSum += amount;
      }
      const currSum = groupSummary.CategorialSum.get(expCatName) || 0;
      groupSummary.CategorialSum.set(expCatName, currSum + amount);
    }
  }

  _populateSummaryData(sheetNames) {
    for(const sheetName of sheetNames) {
      Logger.log(`Generating summary for ${sheetName}`);
      if(this._isNonExpenseSheet(sheetName)) {
        throw new Error(`Cannot generate summary for sheet '${sheetName}'`);
      }
      let sheet = this._spreadSheet.getSheetByName(sheetName);
      this._populateSummaryDataFromSheet(sheet);
    }
  }
  
  _getDataValidationRules() {
    const ruleExpCat = SpreadsheetApp.newDataValidation()
                              .requireValueInList(this._expenseCategories)
                              .setAllowInvalid(false)
                              .setHelpText("Choose from the Dropdown")
                              .build();
    const ruleExpGrp = SpreadsheetApp.newDataValidation()
                              .requireValueInList(this._expenseGroups)
                              .setAllowInvalid(false)
                              .setHelpText("Choose from the Dropdown")
                              .build();
    const rulePayer = SpreadsheetApp.newDataValidation()
                              .requireValueInList(this._payers)
                              .setAllowInvalid(false)
                              .setHelpText("Choose from the Dropdown")
                              .build();
    const rulePayMode = SpreadsheetApp.newDataValidation()
                              .requireValueInList(this._paymentModes)
                              .setAllowInvalid(false)
                              .setHelpText("Choose from the Dropdown")
                              .build();
    return [ruleExpCat, ruleExpGrp, rulePayer, rulePayMode];
  }

  _isNonExpenseSheet(currSheetName) {
    return this._NonExpenseSheetNames.has(currSheetName);
  }

  _computeAllExpenseSheetNames() {
    const sheets = this._spreadSheet.getSheets();
    for(const sheet of sheets) {
      const currSheetName = sheet.getName();
      if(this._isNonExpenseSheet(currSheetName)) {
        continue;
      }
      this._expenseSheetNames.push(currSheetName);
    }
    Logger.log(`Expense Sheet Names: ${this._expenseSheetNames}`);
  }

  _getColumnIndexBasedOnColumnName(headerRow, columnName) {
    const index = headerRow.indexOf(columnName);
    if(index < 0) {
      throw new Error(`${columnName} does not exist!`);
    }
    return index;
  }

  _getSelectionsSheet() {
    let selectionsSheet = this._spreadSheet.getSheetByName(this._selectionsSheetName);
    if(!selectionsSheet) {
      selectionsSheet = this._createSelectionsSheet();
    }
    return selectionsSheet;
  }

  _checkForDuplicatesInArray(arr) {
    if(!Array.isArray(arr)) {
      return;
    }
    const uniqueElems = new Set(arr);
    if(uniqueElems.size === arr.length) {
      return;
    }
    throw new Error("Each column in Selections sheet should be unique");
  }

  _checkForDuplicatesInSelection() {
    this._checkForDuplicatesInArray(this._expenseCategories);
    this._checkForDuplicatesInArray(this._expenseGroups);
    this._checkForDuplicatesInArray(this._payers);
    this._checkForDuplicatesInArray(this._paymentModes);
  }
  
  _getColumnIndices(headerRow) {
    return [
      this._getColumnIndexBasedOnColumnName(headerRow, this._expCatColName),
      this._getColumnIndexBasedOnColumnName(headerRow, this._expGrpColName),
      this._getColumnIndexBasedOnColumnName(headerRow, this._payerColName),
      this._getColumnIndexBasedOnColumnName(headerRow, this._payModeColName)
    ];
  }

  _fetchAllSelections() {
    const selectionSheet = this._getSelectionsSheet();
    const selectionsData = selectionSheet.getDataRange().getValues();
    if(selectionsData.length < 1) {
      Logger.log("Selections sheet is empty or has only header row");
      return;
    }
    const headerRow = selectionsData[0];
    const [expenseCategoryColIdx, expenseGroupColIdx, payerColIdx, paymentModeColIdx] = this._getColumnIndices(headerRow);

    for(let i = 1; i < selectionsData.length; ++i) {
      // start for index 1 as 0th index contains the header
      const currRow = selectionsData[i];
      if(currRow[expenseCategoryColIdx]) {
        this._expenseCategories.push(currRow[expenseCategoryColIdx].toString().trim());
      }
      if(currRow[expenseGroupColIdx]) {
        this._expenseGroups.push(currRow[expenseGroupColIdx].toString().trim());
      }
      if(currRow[payerColIdx]) {
        this._payers.push(currRow[payerColIdx].toString().trim());
      }
      if(currRow[paymentModeColIdx]) {
        this._paymentModes.push(currRow[paymentModeColIdx].toString().trim());
      }
    }
    this._checkForDuplicatesInSelection();

    // Store in set for faster lookup
    this._expenseCategoriesSet = new Set(this._expenseCategories);
    this._expenseGroupsSet = new Set(this._expenseGroups);
    this._payersSet = new Set(this._payers);
    this._paymentModesSet = new Set(this._paymentModes);

    Logger.log(`Found following expense categories ${this._expenseCategories}`);
  }

  _validateSelectionsInSheet(sheetName) {
    Logger.log(`Validating selections in sheet ${sheetName}`);
    const expenseSheet = this._spreadSheet.getSheetByName(sheetName);
    if(!expenseSheet) {
      throw new Error(`${sheetName} does not exist!`);
    }
    const dataRange = expenseSheet.getDataRange();
    const values = dataRange.getValues();
    if(values.length < 1) {
      Logger.log(`${sheetName} sheet is empty or has only header row`);
      return;
    }
    const headerRow = values[0];
    const [expenseCategoryColIdx, expenseGroupColIdx, payerColIdx, paymentModeColIdx]  = this._getColumnIndices(headerRow);
    const amountIdx = this._getColumnIndexBasedOnColumnName(headerRow, this._amountColName);
    // Define column letters for easier highlighting
    const colLetterExpenseCategory = String.fromCharCode(65 + expenseCategoryColIdx);
    const colLetterExpenseGroup = String.fromCharCode(65 + expenseGroupColIdx);
    const colLetterPayer = String.fromCharCode(65 + payerColIdx);
    const colLetterPaymentMode = String.fromCharCode(65 + paymentModeColIdx);
    const colLetterAmount = String.fromCharCode(65 + amountIdx);

    expenseSheet.getDataRange().setBackground(null);
    let invalidCount = 0;
    for(let iRow = 1; iRow < values.length; ++iRow) {
      // start for index 1 as 0th index contains the header
      const row = values[iRow];
      // Spreadsheet user visible row number
      const rowNum = iRow + 1;

      const expenseCategory = row[expenseCategoryColIdx] ? row[expenseCategoryColIdx].toString().trim() : "";
      const expenseGroup = row[expenseGroupColIdx] ? row[expenseGroupColIdx].toString().trim() : "";
      const payer = row[payerColIdx] ? row[payerColIdx].toString().trim() : "";
      const paymentMode = row[paymentModeColIdx] ? row[paymentModeColIdx].toString().trim() : "";

      if(!this._expenseCategoriesSet.has(expenseCategory)) {
        Logger.log(`Invalid Expense Category: '${expenseCategory}' in ${sheetName} row ${rowNum}`);
        expenseSheet.getRange(`${colLetterExpenseCategory}${rowNum}`).setBackground("red");
        ++invalidCount;
      }
      if(!this._expenseGroupsSet.has(expenseGroup)) {
        Logger.log(`Invalid Expense Group: '${expenseGroup}' in ${sheetName} row ${rowNum}`);
        expenseSheet.getRange(`${colLetterExpenseGroup}${rowNum}`).setBackground("red");
        ++invalidCount;
      }
      if(!this._payersSet.has(payer)) {
        Logger.log(`Invalid Payer: '${payer}' in ${sheetName} row ${rowNum}`);
        expenseSheet.getRange(`${colLetterPayer}${rowNum}`).setBackground("red");
        ++invalidCount;
      }
      if(!this._paymentModesSet.has(paymentMode)) {
        Logger.log(`Invalid Payment Mode: '${paymentMode}' in ${sheetName} row ${rowNum}`);
        expenseSheet.getRange(`${colLetterPaymentMode}${rowNum}`).setBackground("red");
        ++invalidCount;
      }
      const amtVal = values[iRow][amountIdx];
      const numericVal = parseFloat(amtVal);
      if(isNaN(numericVal)) {
        Logger.log(`Invalid Amount: '${numericVal}' in ${sheetName} row ${rowNum}`);
        expenseSheet.getRange(`${colLetterAmount}${rowNum}`).setBackground("red");
        throw new Error("Amount should be a number!");
      }
      const amtColNum = amountIdx + 1;
      expenseSheet.getRange(rowNum, amtColNum).setValue(numericVal);
    }
    if (invalidCount === 0) {
      return;
    }
    this._validationErrors.push(`Found ${invalidCount} invalid selections in ${sheetName}`);
  }
  /********** End: Private functions **********/
} // class ExpenseManagerApp
