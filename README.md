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

## Executing the Script with Python

In addition to running the script from the Google Sheets editor, you can execute it from your local machine using the provided Python script. This method is ideal for automated workflows.

### Prerequisites

- Python 3.6 or higher
- Access to a Google Cloud Platform (GCP) project

### 1. Set Up Your Python Environment

- Clone this repository to your local machine.
- Install the required Python libraries using pip:

  ```bash
  pip install -r requirements.txt
  ```

### 2. Configure Google API Access

To run the Python script, you need to enable the Google Apps Script API and create OAuth 2.0 credentials.

- **Enable the Google Apps Script API:**
  - Go to the [Google Cloud Console](https://console.cloud.google.com/).
  - Make sure you have a GCP project created.
  - In the navigation menu, go to `APIs & Services` > `Library`.
  - Search for "Google Apps Script API" and enable it.

- **Create OAuth 2.0 Credentials:**
  - In the navigation menu, go to `APIs & Services` > `Credentials`.
  - Click `+ CREATE CREDENTIALS` and select `OAuth client ID`.
  - Choose `Desktop app` as the application type.
  - Give it a name (e.g., "Fuel Management Script Client").
  - After creating the client ID, a pop-up will show your client ID and secret. Click `DOWNLOAD JSON`.
  - Rename the downloaded file to `creds.json` and place it in the root directory of this project. **Do not commit this file to version control.**

### 3. Link the GCP Project to Your Apps Script

Your Apps Script project needs to be associated with the GCP project where you enabled the API.

- Open your Apps Script project in the editor (`Extensions` > `Apps Script`).
- In the left sidebar, click `Project Settings` (the gear icon).
- Scroll down to the `Google Cloud Platform (GCP) Project` section.
- Click `Change project` and enter the **Project Number** of your GCP project. You can find this number on your GCP Console dashboard.

### 4. Run the Python Script

- Before running the script, make sure you have the `Script_ID` environment variable set. You can find the Script ID in your Apps Script `Project Settings`.

  - **For Linux/macOS:**
    ```bash
    export Script_ID="YOUR_SCRIPT_ID"
    ```
  - **For Windows:**
    ```bash
    set Script_ID="YOUR_SCRIPT_ID"
    ```

- Execute the script:
  ```bash
  python execute_script.py
  ```

- The first time you run it, a browser window will open, asking you to authorize the script. Log in with your Google account and grant the necessary permissions.
- After authorization, a `token.json` file will be created in the project directory. This file stores your credentials, so you won't need to authorize again unless you delete it.
- The script will then call the `createFuelManagementTemplate` function, and you will see a success message in your terminal.