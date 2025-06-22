function onOpen() {
  try {
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Expense Manager");
    menu.addItem("Create new expense sheet", "createExpenseSheet");
    menu.addItem("Validate & setup dropdowns", "setup");
    menu.addItem("Generate summary", "summarize")
    menu.addItem("Monthly report from 'Daily Expenses'", "monthlyReport")
    menu.addToUi();
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

function setup() {
  try {
    let expenseMgrApp = new ExpenseManagerApp();
    expenseMgrApp.setup();
    Browser.msgBox("Success", "Setup completed!", Browser.Buttons.OK);
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

function summarize() {
  try {
    let expenseMgrApp = new ExpenseManagerApp();
    expenseMgrApp.setup();
    expenseMgrApp.generateSummary();
    Browser.msgBox("Success", "Summary generation completed!", Browser.Buttons.OK);
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

function monthlyReport() {
  try {
    let expenseMgrApp = new ExpenseManagerApp();
    expenseMgrApp.setup();
    expenseMgrApp.monthlyReport();
    Browser.msgBox("Success", "Monthly report generation completed!", Browser.Buttons.OK);
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

function createExpenseSheet() {
  try {
    let expenseMgrApp = new ExpenseManagerApp();
    expenseMgrApp.createExpenseSheet();
    Browser.msgBox("Success", "Monthly report generation completed!", Browser.Buttons.OK);
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

class ExpenseManagerApp {
  constructor() {
    this._init();
  }

  setup() {
    this._applyStandardFormatting();
    this._validateSelections();
    this._standardizeDateFormat();
    this._sortByDate();
    this._applyNumericValidation();
    this._setupDropdowns();
  }

  generateSummary() {
    this._populateSummaryData();
    this._generateSummaryImpl();
  }

  monthlyReport() {
    this._populateMonthlyReportData();
    this._generateMonthlyReport();
  }

  createExpenseSheet() {
    const date = new Date();
    const timeZone = this._spreadSheet.getSpreadsheetTimeZone();
    const baseSheetName = Utilities.formatDate(date, timeZone, "MMM yyyy"); // e.g. "Jul 2025"
    let newSheetName = baseSheetName;
    let counter = 0;
    while(this._spreadSheet.getSheetByName(newSheetName) !== null) {
      ++counter;
      newSheetName = `${baseSheetName} (${counter})`; // e.g. "Jul 2025 (1)"
    }
    let newSheet = this._spreadSheet.insertSheet(newSheetName);
    const headerRow = ["Date", "Amount (INR)", "Expense Category", "Expense Description", "Expense Group", "Payer", "Payment Mode"];
    let range = newSheet.getRange(1, 1, 1, headerRow.length);
    range.setValues([headerRow]);
    range.setFontWeight("bold");

    this._applyStandardFormattingToSheet(newSheet);
    this._applyDatePickerToFirstColumnToSheet(newSheet);
    this._applyNumericValidationToSheet(newSheet);
    this._setupDrowdownsToSheet(newSheet);

    Logger.log(`Created new expense sheet ${newSheet.getSheetName()}`);
  }

  _init() {
    this._spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    this._selectionsSheetName = "Selections";
    this._summarySheetName = "Summary";
    this._monthlyReportSheetName = "Monthly Breakdown"
    this._expenseSheetNames = [];
    this._expenseCategories = [];
    this._expenseGroups = [];
    this._payers = [];
    this._paymentModes = [];
    this._validationErrors = [];
    this._summaryData = new Map();
    this._totalSum = 0;
    this._monthlyData = new Map();

    this._computeAllExpenseSheetNames();
    this._fetchAllSelections();
  }

  _applyStandardFormattingToSheet(sheet) {
    Logger.log(`Apply formatting to sheet ${sheet.getName()}`);
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    const lastCol = sheet.getLastColumn() > 0 ? sheet.getLastColumn() : 1;
    let range = sheet.getRange(1, 1, maxRows, maxCols);
    range.setFontFamily("Consolas");
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
    dateRange.setNumberFormat('dd-MMM-yyyy');
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
      const amountIdx = this._getColumnIndexBasedOnColumnName(headerRow, "Amount (INR)");
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
      const dateIdx = this._getColumnIndexBasedOnColumnName(headerRow, "Date");
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
        if(date.toString() !== "Invalid Date") {
          const timezone = this._spreadSheet.getSpreadsheetTimeZone();
          const formattedDate = Utilities.formatDate(date, timezone, "dd-MMM-yyyy");
          newDateValues.push([formattedDate]);
        } else {
          Logger.log(`Row ${i + 2}: Invalid date format`);
          ++invalidDateCount;
          newDateValues.push([originalValue]);
          invalidDateRowNums.push(i + 2);
        }
      }

      dataRange.setValues(newDateValues);
      dataRange.setNumberFormat("dd-MMM-yyyy");
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

  _populateMonthlyReportData() {
    for(const sheetName of this._expenseSheetNames) {
      // Iterate through all sheets; multiple sheets might contain "Daily Expenses" group data.
      Logger.log(`Analyzing ${sheetName} for computing Monthly report from Daily Expenses`);
      let sheet = this._spreadSheet.getSheetByName(sheetName);
      const values = sheet.getDataRange().getValues();
      const headerRow = sheet.getDataRange().getValues()[0];
      const amountIdx = this._getColumnIndexBasedOnColumnName(headerRow, "Amount (INR)");
      const expCatIdx = this._getColumnIndexBasedOnColumnName(headerRow, "Expense Category");
      const expGrpIdx = this._getColumnIndexBasedOnColumnName(headerRow, "Expense Group");
      const dateIdx = this._getColumnIndexBasedOnColumnName(headerRow, "Date");

      for(let i = 1; i < values.length; ++i) {
        // start for index 1 as 0th index contains the header
        const currRow = values[i];

        const expGrpName = currRow[expGrpIdx].toString().trim();
        if(expGrpName !== "Daily Expenses") {
          continue;
        }
        const expCatName = currRow[expCatIdx].toString().trim();
        const amount = parseFloat(currRow[amountIdx]);
        const dateStr = currRow[dateIdx].toString().trim();
        const dateObj = new Date(dateStr);
        const monthYearStr = this._getMonthYearString(dateObj);

        let monthlySummary = this._monthlyData.get(monthYearStr);
        if(!monthlySummary) {
          monthlySummary = {"MonthlyTotal": 0.0, "CategoricalSum": new Map()};
          this._monthlyData.set(monthYearStr, monthlySummary);
        }
        if(expCatName !== "Refund") {
          monthlySummary.MonthlyTotal += amount;
        }
        const currSum = monthlySummary.CategoricalSum.get(expCatName) || 0;
        monthlySummary.CategoricalSum.set(expCatName, currSum + amount);
      }
    }
  }

  _generateMonthlyReport() {
    let output = [];
    let headerRow = ["Month", ...this._expenseCategories, "Monthly Total (INR)"];
    output.push(headerRow);

    for(const [monthYearStr, monthInfo] of this._monthlyData) {
      let monthRow = [];
      monthRow.push(monthYearStr);
      for(const categoryName of this._expenseCategories) {
        const categorySumForTheMonth = monthInfo.CategoricalSum.get(categoryName) || 0;
        monthRow.push(categorySumForTheMonth);
      }
      monthRow.push(monthInfo.MonthlyTotal);
      output.push(monthRow);
    }
    let monthlyReportSheet = this._spreadSheet.getSheetByName(this._monthlyReportSheetName);
    if(monthlyReportSheet) {
      monthlyReportSheet.clearContents();
    } else {
      monthlyReportSheet = this._spreadSheet.insertSheet(this._monthlyReportSheetName);
    }

    let range = monthlyReportSheet.getRange(1, 1, output.length, output[0].length);
    range.setValues(output);

    // Bold header row
    monthlyReportSheet.getRange(1, 1, 1, output[0].length).setFontWeight("bold");
    Logger.log(`Monthly report Generated in ${this._summarySheetName} sheet`);
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

  _generateSummaryImpl() {
    let output = [];
    output.push(['Expense Group', 'Expense Category', 'Amount (INR)', '% Within Group', '% of Overall Total']);
    for(const [groupName, groupInfo] of this._summaryData) {
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
    let summarySheet = this._spreadSheet.getSheetByName(this._summarySheetName);
    if(summarySheet) {
      summarySheet.clearContents();
    } else {
      summarySheet = this._spreadSheet.insertSheet(this._summarySheetName);
    }

    let range = summarySheet.getRange(1, 1, output.length, output[0].length);
    range.setValues(output);

    summarySheet.getRange(1, 1, 1, output[0].length).setFontWeight("bold");
    Logger.log(`Expense Summary Generated in ${this._summarySheetName} sheet`);
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

  _populateSummaryData() {
    for(const sheetName of this._expenseSheetNames) {
      Logger.log(`Generating summary for ${sheetName}`);
      let sheet = this._spreadSheet.getSheetByName(sheetName);
      const values = sheet.getDataRange().getValues();
      const headerRow = values[0];
      const amountIdx = this._getColumnIndexBasedOnColumnName(headerRow, "Amount (INR)");
      const expCatIdx = this._getColumnIndexBasedOnColumnName(headerRow, "Expense Category");
      const expGrpIdx = this._getColumnIndexBasedOnColumnName(headerRow, "Expense Group");

      for(let i = 1; i < values.length; ++i) {
        // start for index 1 as 0th index contains the header
        const currRow = values[i];

        const expGrpName = currRow[expGrpIdx].toString().trim();
        const expCatName = currRow[expCatIdx].toString().trim();
        const amount = parseFloat(currRow[amountIdx]);

        let groupSummary = this._summaryData.get(expGrpName);
        if(!groupSummary) {
          groupSummary = {"GroupTotal": 0.0, "CategorialSum": new Map()};
          this._summaryData.set(expGrpName, groupSummary);
        }
        if(expCatName !== "Refund") {
          groupSummary.GroupTotal += amount;
          this._totalSum += amount;
        }
        const currSum = groupSummary.CategorialSum.get(expCatName) || 0;
        groupSummary.CategorialSum.set(expCatName, currSum + amount);
      }
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

  _isExcludedSheet(currSheetName) {
    return currSheetName === this._selectionsSheetName 
    || currSheetName === this._summarySheetName 
    || currSheetName === this._monthlyReportSheetName;
  }

  _computeAllExpenseSheetNames() {
    const sheets = this._spreadSheet.getSheets();
    for(const sheet of sheets) {
      const currSheetName = sheet.getName();
      if(this._isExcludedSheet(currSheetName)) {
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
    const selectionsSheet = this._spreadSheet.getSheetByName(this._selectionsSheetName);
    if(!selectionsSheet) {
      throw new Error(`${this._selectionsSheetName} does not exist!`);
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
      this._getColumnIndexBasedOnColumnName(headerRow, "Expense Category"),
      this._getColumnIndexBasedOnColumnName(headerRow, "Expense Group"),
      this._getColumnIndexBasedOnColumnName(headerRow, "Payer"),
      this._getColumnIndexBasedOnColumnName(headerRow, "Payment Mode")
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
    const amountIdx = this._getColumnIndexBasedOnColumnName(headerRow, "Amount (INR)");
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

      if(!this._expenseCategories.includes(expenseCategory)) {
        Logger.log(`Invalid Expense Category: '${expenseCategory}' in ${sheetName} row ${rowNum}`);
        expenseSheet.getRange(`${colLetterExpenseCategory}${rowNum}`).setBackground("red");
        ++invalidCount;
      }
      if(!this._expenseGroups.includes(expenseGroup)) {
        Logger.log(`Invalid Expense Group: '${expenseGroup}' in ${sheetName} row ${rowNum}`);
        expenseSheet.getRange(`${colLetterExpenseGroup}${rowNum}`).setBackground("red");
        ++invalidCount;
      }
      if(!this._payers.includes(payer)) {
        Logger.log(`Invalid Payer: '${payer}' in ${sheetName} row ${rowNum}`);
        expenseSheet.getRange(`${colLetterPayer}${rowNum}`).setBackground("red");
        ++invalidCount;
      }
      if(!this._paymentModes.includes(paymentMode)) {
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
    if (invalidCount > 0) {
      this._validationErrors.push(`Found ${invalidCount} invalid selections in ${sheetName}`);
    }
  }
}
