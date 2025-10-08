/**
 * @OnlyCurrentDoc
 *
 * The above comment directs App Script to limit the scope of script authorization
 * to the current spreadsheet only. It's a best practice for security and
 * user trust.
 */

/**
 * Creates a complete fuel management template in the active spreadsheet.
 *
 * This function serves as the main entry point for the script. It orchestrates
 * the creation of the entire template by calling a series of helper functions,
 * each responsible for a specific sheet. It concludes by flushing pending
 * changes and displaying a success alert.
 *
 * @returns {void}
 */
function createFuelManagementTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  clearAllSheets(ss);
  createSettingsSheet(ss);
  createFuelPurchasesSheet(ss);
  createFuelSalesSheet(ss);
  createDashboardSheet(ss);
  SpreadsheetApp.flush(); // Applies all pending spreadsheet changes
  SpreadsheetApp.getUi().alert('Fuel Management Template created successfully!');
}

/**
 * Deletes all existing sheets in the spreadsheet.
 *
 * This function iterates through all sheets in the provided spreadsheet and
 * deletes them, ensuring a clean slate before creating the new template.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet instance.
 * @returns {void}
 */
function clearAllSheets(ss) {
  ss.getSheets().forEach(sheet => ss.deleteSheet(sheet));
}

/**
 * Creates the 'Settings' sheet with predefined parameters for fuel calculations.
 *
 * This sheet holds key configuration values, such as conversion factors and
 * default profit margins, which are referenced in formulas across other sheets.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet instance.
 * @returns {void}
 */
function createSettingsSheet(ss) {
  const sheet = ss.insertSheet('Settings');
  // Set headers and initial values for fuel parameters
  sheet.getRange('A1:B1').setValues([['Parameter', 'Gasoline']]);
  sheet.getRange('A2:A6').setValues([
    ['Liters per Ton'],
    ['Liters per Tank'],
    ['Default Profit per Liter (៛)'],
    ['Default Profit per Tank (៛)'],
    ['Default Exchange Rate USD→KHR']
  ]);
  sheet.getRange('B2:B5').setValues([[1390], [34], [300], [4000]]);
  sheet.getRange('B6').setValue(4100); // Placeholder for exchange rate
  sheet.getRange('C1').setValue('Diesel');
  sheet.getRange('C2:C5').setValues([[1190], [34], [300], [4000]]);
  sheet.getRange('C6').setValue(4100); // Placeholder for exchange rate
}

/**
 * Creates the 'Fuel Purchases' sheet for tracking fuel acquisition data.
 *
 * This sheet includes headers for manual data entry (e.g., date, price) and
 * columns with pre-filled formulas to automatically calculate metrics like
 * total cost, cost per liter, and cost per tank. It also adds data validation
 * to guide user input.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet instance.
 * @returns {void}
 */
function createFuelPurchasesSheet(ss) {
  const sheet = ss.insertSheet('Fuel Purchases');
  const headers = [
    'Date', 'Fuel Type', 'Buy Price per Ton', 'Currency', 'Amount Bought (Ton)',
    'Custom Exchange USD→KHR', 'Custom Exchange KHR→USD',
    'Total Buy (USD)', 'Total Buy (KHR)', 'Liters', 'Tanks',
    'Buy per Liter (USD)', 'Buy per Liter (KHR)', 'Buy per Tank (USD)', 'Buy per Tank (KHR)'
  ];
  sheet.getRange('A1:O1').setValues([headers]);

  // --- Data Validation Rules ---
  // Dropdown for 'Fuel Type' column
  const fuelTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Gasoline', 'Diesel'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B2:B').setDataValidation(fuelTypeRule);

  // Dropdown for 'Currency' column
  const currencyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['USD', 'KHR'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('D2:D').setDataValidation(currencyRule);

  // --- Formulas with Error Handling ---
  const formulas = [
    '=IFERROR(IF($D2="USD", $C2*$E2, $C2*$E2/$G2), "")',
    '=IFERROR(IF($D2="KHR", $C2*$E2, $C2*$E2*$F2), "")',
    '=IFERROR($E2*IF($B2="Gasoline", Settings!$B$2, Settings!$C$2), "")',
    '=IFERROR(J2/IF($B2="Gasoline", Settings!$B$3, Settings!$C$3), "")',
    '=IFERROR($H2/J2, "")',
    '=IFERROR($I2/J2, "")',
    '=IFERROR(L2*IF($B2="Gasoline", Settings!$B$3, Settings!$C$3), "")',
    '=IFERROR(M2*IF($B2="Gasoline", Settings!$B$3, Settings!$C$3), "")'
  ];
  sheet.getRange('H2:O2').setFormulas([formulas]);
}

/**
 * Creates the 'Fuel Sales' sheet for tracking fuel sales and profitability.
 *
 * This sheet allows users to input sales data and automatically calculates
 * revenue and profit metrics by looking up purchase prices and applying
 * user-defined profit margins. Formulas are wrapped in IFERROR to handle
 * cases where purchase data is missing.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet instance.
 * @returns {void}
 */
function createFuelSalesSheet(ss) {
  const sheet = ss.insertSheet('Fuel Sales');
  const headers = [
    'Date', 'Fuel Type', 'Number of Tanks Sold', 'Profit per Tank (៛)',
    'Sell per Tank (KHR)', 'Sell per Liter (KHR)', 'Total Revenue (KHR)',
    'Total Revenue (USD)', 'Total Profit (KHR)', 'Total Profit (USD)'
  ];
  sheet.getRange('A1:J1').setValues([headers]);

  // --- Data Validation Rules ---
  // Dropdown for 'Fuel Type' column
  const fuelTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Gasoline', 'Diesel'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B2:B').setDataValidation(fuelTypeRule);

  // --- Formulas with Error Handling ---
  const vlookupBase = "VLOOKUP($B2, 'Fuel Purchases'!$B:$O, 15, FALSE)";
  const formulas = [
    `=IFERROR(${vlookupBase} + $D2, "Missing Purchase Data")`,
    `=IFERROR(E2 / IF($B2="Gasoline", Settings!$B$3, Settings!$C$3), "")`,
    `=IFERROR(E2 * $C2, "")`,
    `=IFERROR(G2 / INDIRECT("'Fuel Purchases'!F" & MATCH($B2,'Fuel Purchases'!$B:$B,0)), "")`,
    `=IFERROR(G2 - (${vlookupBase} * $C2), "")`,
    `=IFERROR(I2 / INDIRECT("'Fuel Purchases'!G" & MATCH($B2,'Fuel Purchases'!$B:$B,0)), "")`
  ];
  sheet.getRange('E2:J2').setFormulas([formulas]);
}

/**
 * Creates the 'Dashboard' sheet to summarize key financial metrics.
 *
 * This sheet provides a high-level overview of the fuel business by using
 * SUMIF formulas to aggregate data from the 'Fuel Purchases' and 'Fuel Sales'
 * sheets, breaking down totals by fuel type and currency.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet instance.
 * @returns {void}
 */
function createDashboardSheet(ss) {
  const sheet = ss.insertSheet('Dashboard');
  sheet.getRange('A1:B1').setValues([['Metric', 'Value']]);
  const metrics = [
    'Total Gasoline Bought USD', 'Total Diesel Bought USD',
    'Total Gasoline Bought KHR', 'Total Diesel Bought KHR',
    'Total Revenue USD', 'Total Revenue KHR',
    'Total Profit USD', 'Total Profit KHR'
  ];
  sheet.getRange('A2:A9').setValues(metrics.map(m => [m]));

  // Set formulas to aggregate data from other sheets.
  sheet.getRange('B2').setFormula("=SUMIF('Fuel Purchases'!B:B,\"Gasoline\",'Fuel Purchases'!H:H)");
  sheet.getRange('B3').setFormula("=SUMIF('Fuel Purchases'!B:B,\"Diesel\",'Fuel Purchases'!H:H)");
  sheet.getRange('B4').setFormula("=SUMIF('Fuel Purchases'!B:B,\"Gasoline\",'Fuel Purchases'!I:I)");
  sheet.getRange('B5').setFormula("=SUMIF('Fuel Purchases'!B:B,\"Diesel\",'Fuel Purchases'!I:I)");
  sheet.getRange('B6').setFormula("=SUM('Fuel Sales'!H:H)");
  sheet.getRange('B7').setFormula("=SUM('Fuel Sales'!G:G)");
  sheet.getRange('B8').setFormula("=SUM('Fuel Sales'!J:J)");
  sheet.getRange('B9').setFormula("=SUM('Fuel Sales'!I:I)");
}