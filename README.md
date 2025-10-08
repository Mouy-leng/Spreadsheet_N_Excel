# Fuel Management Spreadsheet Automation

This repository provides a Google Apps Script to automatically generate a comprehensive fuel management spreadsheet. The script is designed to streamline fuel tracking for businesses by creating a ready-to-use template with automated calculations.

It can be executed directly from Google Sheets or run remotely via a Python script, making it suitable for automated workflows.

## Features

- **Automated Setup**: Instantly creates a multi-sheet spreadsheet.
- **Four-Sheet System**:
    - **Settings**: Configure default values for fuel types (Gasoline, Diesel), such as liters per ton, liters per tank, and default profit margins.
    - **Fuel Purchases**: Track incoming fuel purchases, with automatic calculations for total cost, cost per liter, and cost per tank in both USD and KHR.
    - **Fuel Sales**: Record fuel sales and automatically calculate revenue and profit based on purchase data and custom profit margins.
    - **Dashboard**: Get a high-level financial overview with key metrics like total purchases, total revenue, and total profit, all summarized automatically.
- **Dual Currency Support**: Formulas are pre-configured to handle transactions in both USD and KHR.
- **Remote Execution**: Includes a Python script to run the template generation from any local environment, enabling automation.

## Getting Started

### Prerequisites

- A Google Account.
- For local execution:
    - Python 3.6+
    - A Google Cloud Platform (GCP) project.

---

### Method 1: Deploying from Google Sheets

This is the simplest way to get started.

1.  **Create a Google Sheet**: Go to [sheets.google.com](https://sheets.google.com) and create a new, blank spreadsheet.
2.  **Open the Apps Script Editor**: In your spreadsheet, navigate to `Extensions` > `Apps Script`.
3.  **Add the Script Code**:
    - Delete any default code in the `Code.gs` file.
    - Copy the entire content of `Code.gs` from this repository and paste it into the editor.
4.  **Save and Run**:
    - Click the **Save project** icon.
    - Ensure `createFuelManagementTemplate` is the selected function in the toolbar.
    - Click **Run**.
5.  **Grant Permissions**: The first time you run the script, Google will ask for permission to access your spreadsheet. Follow the prompts to authorize it. You may need to click "Advanced" and "Go to (your project name)" to proceed.
6.  **Done!**: A success message will appear, and your spreadsheet will now contain the `Settings`, `Fuel Purchases`, `Fuel Sales`, and `Dashboard` sheets.

---

### Method 2: Executing Remotely with Python

This method is ideal for automated workflows or for developers who prefer working from a local environment.

#### 1. Set Up Your Python Environment

- Clone this repository to your local machine.
- Install the required Python libraries:
  ```bash
  pip install -r requirements.txt
  ```

#### 2. Configure Google API Access

To run the Python script, you need to enable the Google Apps Script API and create OAuth 2.0 credentials for a desktop application.

- **Enable the API**:
    1.  Go to the [Google Cloud Console](https://console.cloud.google.com/).
    2.  Select your GCP project.
    3.  Navigate to `APIs & Services` > `Library`.
    4.  Search for "Google Apps Script API" and **enable** it.

- **Create OAuth 2.0 Credentials**:
    1.  Go to `APIs & Services` > `Credentials`.
    2.  Click `+ CREATE CREDENTIALS` and select `OAuth client ID`.
    3.  Choose `Desktop app` as the application type.
    4.  After creation, click `DOWNLOAD JSON` and save the file as `creds.json` in the root of this project. **Warning: Do not commit `creds.json` to version control.**

#### 3. Link Your GCP Project to Apps Script

The Apps Script project inside your Google Sheet needs to be linked to the GCP project where you enabled the API.

1.  Open your Apps Script project again (`Extensions` > `Apps Script`).
2.  In the left sidebar, click **Project Settings** (⚙️ icon).
3.  Scroll to the `Google Cloud Platform (GCP) Project` section.
4.  Click `Change project` and enter the **Project Number** of your GCP project (found on your GCP Console dashboard).

#### 4. Run the Python Script

1.  **Find Your Script ID**: Open your Apps Script project and go to **Project Settings** (⚙️ icon). Copy the **Script ID**.

2.  **Execute the script**: Run the following command in your terminal, replacing `YOUR_SCRIPT_ID` with the ID you copied.
    ```bash
    python execute_script.py YOUR_SCRIPT_ID
    ```

3.  **Authorize**: The first time you run the script, a browser window will open for you to log in and grant permissions. A `token.json` file will be created to store your credentials for future runs.

Upon successful execution, the script will print a success message to your terminal.

## Sheet Descriptions

-   **`Settings`**: Holds global variables for fuel calculations (e.g., Liters per Ton, default profit margins).
-   **`Fuel Purchases`**: Log all fuel buys here. Columns with light gray headers are auto-calculated.
-   **`Fuel Sales`**: Log all fuel sales here. Columns with light gray headers are auto-calculated.
-   **`Dashboard`**: A read-only sheet that provides a summary of your financial data.