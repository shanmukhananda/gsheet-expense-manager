function onOpen() {
  try {
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Expense Manager");
    menu.addItem("Validate & Setup Dropdowns", "setup");
    menu.addItem("Generate Summary", "summarize")
    menu.addToUi();
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

function setup() {
  try {
    let expenseMgrApp = new ExpenseManagerApp();
    expenseMgrApp.validateSelections();
    expenseMgrApp.setupDropdowns();

    Browser.msgBox("Success", "Setup Completed!", Browser.Buttons.OK);
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

function summarize() {
  try {
    let expenseMgrApp = new ExpenseManagerApp();
    expenseMgrApp.validateSelections();
    expenseMgrApp.generateSummary();
    Browser.msgBox("Success", "Summary Generation Completed!", Browser.Buttons.OK);
  } catch (err) {
      Logger.log(`Error occured ${err.message}`);
      Browser.msgBox("Error", err.message, Browser.Buttons.OK);
  }
}

class ExpenseManagerApp {
  constructor() {
    this._init();
  }

  _init() {
    this._spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    this._selectionsSheetName = "Selections";
    this._summarySheetName = "Summary";
    this._expenseSheetNames = [];
    this._expenseCategories = [];
    this._expenseGroups = [];
    this._payers = [];
    this._paymentModes = [];
    this._validationErrors = [];
    this._summaryData = new Map();
    this._totalSum = 0;

    this._computeAllExpenseSheetNames();
    this._fetchAllSelections();
  }

  validateSelections() {
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

  setupDropdowns() {
    const [ruleExpCat, ruleExpGrp, rulePayer, rulePayMode] = this._getDataValidationRules();
    for(const sheetName of this._expenseSheetNames) {
      Logger.log(`Adding data validation rule for ${sheetName}`);
      let sheet = this._spreadSheet.getSheetByName(sheetName);
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
  }

  generateSummary() {
    this._populateSummaryData();
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

    // Format
    // Bold header row
    summarySheet.getRange(1, 1, 1, output[0].length).setFontWeight("bold");
    // Auto-resize columns
    summarySheet.autoResizeColumns(1, output[0].length);

    Logger.log(`Expense Summary Generated in ${this._summarySheetName} sheet`);
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

  _computeAllExpenseSheetNames() {
    const sheets = this._spreadSheet.getSheets();
    for(const sheet of sheets) {
      const currSheetName = sheet.getName();
      if(currSheetName === this._selectionsSheetName || currSheetName === this._summarySheetName) {
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
