# Google App Script Expense Manager

This Google App Script simplifies expense tracking in Google Sheets. It automates formatting, data validation, and generates insightful reports like summaries, monthly breakdowns, and payer spending analysis.

## Features

The "Expense Manager" custom menu in Google Sheets provides:

* **Setup**: Initializes spreadsheet, sets up data validation, formats sheets, and creates/updates the "Selections" sheet.

* **Create new expense sheet**: Generates a new expense sheet with pre-configured headers, validation, and dropdowns.

* **Generate Summary**: Creates an expense summary for the current sheet or all sheets.

* **Monthly Report**: Generates a monthly spending report, primarily for "Daily Expenses," for the current sheet or all sheets.

* **Payer spending report**: Provides a breakdown of spending by payer and expense group.

## Installation

1.  Open your Google Sheet and go to `Extensions` > `App Script`.

2.  Copy the `Code.js` content into the `Code.gs` file in the script editor.

3.  Save the script.

4.  Refresh your Google Sheet; the "Expense Manager" menu will appear.

## Setup and Usage

1.  **Run Setup**: Click `Expense Manager` > `Setup`. Authorize the script if prompted. This populates your "Selections" sheet.

2.  **Configure Selections**: Customize "Expense Category," "Expense Group," "Payer," and "Payment Mode" in the "Selections" sheet. Ensure no duplicates.

3.  **Create Sheets**: Use `Expense Manager` > `Create new expense sheet` to add new expense logging sheets.

4.  **Log Expenses**: Enter data into your expense sheets. `Date` (with a date picker), `Amount (INR)` (numeric only), and dropdowns for other categories will guide input.

5.  **Generate Reports**: Use the menu options for "Generate summary," "Monthly report," and "Payer spending report" to analyze your expenses.

## Sheet Structure

* **Expense Sheets**: Any sheet not named "Selections", "Summary", "Monthly Breakdown", or "Payer Spending Report". Must contain: `Date`, `Amount (INR)`, `Expense Category`, `Expense Description`, `Expense Group`, `Payer`, `Payment Mode`.

* **"Selections" Sheet**: Holds master lists for dropdown values.

* **Report Sheets**: `Summary`, `Monthly Breakdown`, `Payer Spending Report` are automatically generated/updated.

## Important Notes

* **Authorization**: Grant permissions when first prompted.

* **Errors**: Red-highlighted cells indicate validation errors. Check App Script `Logs` for details.

* **Customization**: Modify the `_init()` method in `ExpenseManagerApp` for advanced changes.