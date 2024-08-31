function weeklyBalanceCheck() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // First, update transactions
  loadTransactions();

  // Then, check balances
  const accountData = getAccountDataFromSpreadsheet(spreadsheet);
  const accessToken = getAccessToken(); // Implement this function to get the access token

  accountData.forEach(({ accountId, sheetName, customName }) => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Sheet ${sheetName} not found for account ${accountId}`);
      return;
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const balanceColumnIndex = headers.findIndex(header => header === customName) + 1;

    if (balanceColumnIndex === 0) {
      Logger.log(`Balance column not found for account ${accountId} in sheet ${sheetName}`);
      return;
    }

    const lastRow = sheet.getLastRow();
    const calculatedBalance = parseFloat(sheet.getRange(lastRow, balanceColumnIndex).getValue());

    const actualBalance = fetchAccountBalance(accessToken, accountId);

    if (actualBalance !== null) {
      const discrepancy = Math.abs(actualBalance - calculatedBalance);
      if (discrepancy > 0.01) { // Allow for small rounding differences
        Logger.log(`Balance discrepancy found for account ${accountId} (${customName})`);
        Logger.log(`Calculated balance: ${calculatedBalance}, Actual balance: ${actualBalance}`);
        // You might want to highlight this discrepancy in the sheet or send an alert
      } else {
        Logger.log(`Balance check passed for account ${accountId} (${customName})`);
      }
    }
  });
}