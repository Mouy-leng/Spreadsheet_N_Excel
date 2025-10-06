# Google Apps Script for Fuel Management Spreadsheet

This repository contains a Google Apps Script that automatically generates a comprehensive fuel management spreadsheet. The script creates four sheets: `Settings`, `Fuel Purchases`, `Fuel Sales`, and a `Dashboard`, complete with predefined formulas to track and summarize fuel-related data.

## How to Use

Follow these steps to deploy and use the script in your Google Sheets.

### 1. Create a new Google Sheet

- Go to [sheets.google.com](https://sheets.google.com) and create a new, blank spreadsheet.

### 2. Open the Apps Script Editor

- In your new spreadsheet, navigate to `Extensions` > `Apps Script`.
- This will open the script editor in a new tab.

### 3. Add the Script Code

- The script editor will open with a default `Code.gs` file.
- Delete any content in that file.
- Copy the entire content of the `Code.gs` file from this repository and paste it into the script editor.

### 4. Save the Script

- Click the floppy disk icon (Save project) in the script editor's toolbar.
- You can give the project a name, like "Fuel Management".

### 5. Run the Script

- With the `Code.gs` file open, ensure the `createFuelManagementTemplate` function is selected in the dropdown menu next to the "Debug" and "Run" buttons.
- Click the **Run** button.

### 6. Grant Permissions

- The first time you run the script, Google will prompt you to grant permissions for the script to access your spreadsheet.
- Follow the on-screen instructions to authorize the script. You may see a warning that the app isn't verified; you can proceed by clicking "Advanced" and then "Go to (your project name)".

### 7. Check Your Spreadsheet

- Once the script finishes running, you will see an alert saying "Fuel Management Template created successfully!".
- Return to your spreadsheet. You will now see the four new sheets: `Settings`, `Fuel Purchases`, `Fuel Sales`, and `Dashboard`, all populated with headers and formulas.

## How It Works

- **`createFuelManagementTemplate()`**: The main function that runs all the other functions.
- **`clearAllSheets(ss)`**: Deletes any existing sheets in the spreadsheet.
- **`createSettingsSheet(ss)`**: Creates the `Settings` sheet with default values for fuel calculations.
- **`createFuelPurchasesSheet(ss)`**: Creates the `Fuel Purchases` sheet with formulas to calculate purchase metrics.
- **`createFuelSalesSheet(ss)`**: Creates the `Fuel Sales` sheet with formulas to calculate sales and profit.
- **`createDashboardSheet(ss)`**: Creates the `Dashboard` with `SUMIF` formulas to summarize the data from the other sheets.