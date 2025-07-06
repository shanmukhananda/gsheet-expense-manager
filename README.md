# Google Apps Script Expense Manager

A Google Apps Script to simplify expense tracking in Google Sheets. It automates formatting, data validation, and generates insightful reports.

## Features

The "Expense Manager" custom menu provides:

*   **Setup**: Initializes the spreadsheet with required sheets, formatting, and data validation.
*   **Create New Expense Sheet**: Generates a new, pre-configured sheet for logging expenses.
*   **Category-wise Summary**: Breaks down spending by category.
*   **Payer-wise Report**: Analyzes spending by payer and expense group.
*   **Monthly Report**: Generates a monthly spending summary.
*   **Yearly Report**: Generates a yearly spending summary.

## Getting Started

1.  **Open Script Editor**: In your Google Sheet, go to `Extensions` > `Apps Script`.
2.  **Add the Code**: Copy the content from `Code.js` and paste it into the `Code.gs` file in the editor. Save the project.
3.  **Refresh Sheet**: Refresh your Google Sheet. The "Expense Manager" menu will appear.
4.  **Run Setup**: Click `Expense Manager` > `Setup`. Authorize the script when prompted. This will create a "Selections" sheet.
5.  **Customize**: Go to the "Selections" sheet to customize your expense categories, payers, etc.
6.  **Track & Report**:
    *   Use `Expense Manager` > `Create new expense sheet` to add sheets.
    *   Log your expenses in these sheets.
    *   Use the other menu options to generate reports.

## Sheet Structure

*   **Expense Sheets**: Where you log your daily expenses. The script will set up the required columns (`Date`, `Amount (INR)`, `Expense Category`, etc.).
*   **"Selections" Sheet**: Contains the master lists for the dropdowns used in your expense sheets (e.g., categories, payers).
*   **Report Sheets**: Sheets like `Summary`, `Monthly Breakdown`, and `Payer Spending Report` are automatically created and updated by the script when you run a report.

## Notes

*   **Data Validation**: Cells with invalid data (e.g., text in an amount field) will be highlighted in red.
*   **Customization**: For advanced changes, modify the `_init()` method in the script.