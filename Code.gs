/**
 * @OnlyCurrentDoc
 *
 * The above comment directs App Script to limit the scope of script authorization
 * to the current spreadsheet only. It's a best practice for security and
 * user trust.
 */

/**
 * Creates a complete fuel management template in the active spreadsheet.
 * This function acts as the main entry point and orchestrates the creation
 * of the entire template by calling helper functions for each sheet.
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
 * Deletes all existing sheets in the spreadsheet to ensure a clean slate.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 */
function clearAllSheets(ss) {
  ss.getSheets().forEach(sheet => ss.deleteSheet(sheet));
}

/**
 * Creates the 'Settings' sheet with predefined parameters for gasoline and diesel.
 * This sheet holds configuration values used in formulas across other sheets.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
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
  sheet.getRange('B6').setValue(''); // Placeholder for exchange rate
  sheet.getRange('C1').setValue('Diesel');
  sheet.getRange('C2:C5').setValues([[1190], [34], [300], [4000]]);
  sheet.getRange('C6').setValue(''); // Placeholder for exchange rate
}

/**
 * Creates the 'Fuel Purchases' sheet with headers and formulas for tracking fuel buys.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
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

  // Set formulas for calculated columns
  sheet.getRange('H2').setFormula('=IF(D2="USD",C2*E2,C2*E2/G2)');
  sheet.getRange('I2').setFormula('=IF(D2="KHR",C2*E2,C2*E2*F2)');
  sheet.getRange('J2').setFormula('=E2*IF(B2="Gasoline",Settings!B2,Settings!C2)');
  sheet.getRange('K2').setFormula('=J2/IF(B2="Gasoline",Settings!B3,Settings!C3)');
  sheet.getRange('L2').setFormula('=H2/J2');
  sheet.getRange('M2').setFormula('=I2/J2');
  sheet.getRange('N2').setFormula('=L2*IF(B2="Gasoline",Settings!B3,Settings!C3)');
  sheet.getRange('O2').setFormula('=M2*IF(B2="Gasoline",Settings!B3,Settings!C3)');
}

/**
 * Creates the 'Fuel Sales' sheet with headers and formulas for tracking fuel sales.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 */
function createFuelSalesSheet(ss) {
  const sheet = ss.insertSheet('Fuel Sales');
  const headers = [
    'Date', 'Fuel Type', 'Number of Tanks Sold', 'Profit per Tank (៛)',
    'Sell per Tank (KHR)', 'Sell per Liter (KHR)', 'Total Revenue (KHR)',
    'Total Revenue (USD)', 'Total Profit (KHR)', 'Total Profit (USD)'
  ];
  sheet.getRange('A1:J1').setValues([headers]);

  // Set formulas for calculated columns
  sheet.getRange('E2').setFormula("=VLOOKUP(B2,'Fuel Purchases'!B:O,15,FALSE)+D2");
  sheet.getRange('F2').setFormula("=E2/IF(B2=\"Gasoline\",Settings!B3,Settings!C3)");
  sheet.getRange('G2').setFormula("=E2*C2");
  sheet.getRange('H2').setFormula("=G2/'Fuel Purchases'!F2");
  sheet.getRange('I2').setFormula("=G2-(VLOOKUP(B2,'Fuel Purchases'!B:O,15,FALSE)*C2)");
  sheet.getRange('J2').setFormula("=I2/'Fuel Purchases'!G2");
}

/**
 * Creates the 'Dashboard' sheet with formulas to summarize key metrics.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
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

  // Set formulas for dashboard metrics
  sheet.getRange('B2').setFormula("=SUMIF('Fuel Purchases'!B:B,\"Gasoline\",'Fuel Purchases'!H:H)");
  sheet.getRange('B3').setFormula("=SUMIF('Fuel Purchases'!B:B,\"Diesel\",'Fuel Purchases'!H:H)");
  sheet.getRange('B4').setFormula("=SUMIF('Fuel Purchases'!B:B,\"Gasoline\",'Fuel Purchases'!I:I)");
  sheet.getRange('B5').setFormula("=SUMIF('Fuel Purchases'!B:B,\"Diesel\",'Fuel Purchases'!I:I)");
  sheet.getRange('B6').setFormula("=SUM('Fuel Sales'!H:H)");
  sheet.getRange('B7').setFormula("=SUM('Fuel Sales'!G:G)");
  sheet.getRange('B8').setFormula("=SUM('Fuel Sales'!J:J)");
  sheet.getRange('B9').setFormula("=SUM('Fuel Sales'!I:I)");
}