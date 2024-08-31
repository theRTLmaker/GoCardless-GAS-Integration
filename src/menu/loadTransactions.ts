function loadTransactions() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const requisitionsSheet = spreadsheet.getSheetByName(REQUISITIONS_SHEET_NAME);

  if (!requisitionsSheet) {
    throw new Error(`Sheet "${REQUISITIONS_SHEET_NAME}" not found. Please run the initialization first.`);
  }

  scriptLock(_loadTransactions);
}

function _loadTransactions() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const accountData = getAccountDataFromSpreadsheet(spreadsheet);

  if (accountData.length === 0) {
    SpreadsheetApp.getActive().toast("No accounts found. Please link and fetch accounts first.", "Load Transactions");
    return;
  }

  const accessToken = getAccessToken();
  let totalTransactions = 0;
  let processedAccounts = 0;
  let rateLimitedAccounts = 0;

  accountData.forEach(({ accountId, sheetName }) => {
    try {
      const transactions = fetchTransactionsForAccount(accessToken, accountId);
      if (transactions === null) {
        // Rate limit error occurred
        rateLimitedAccounts++;
      } else if (transactions && transactions.length > 0) {
        storeTransactions(spreadsheet, accountId, sheetName, transactions);
        totalTransactions += transactions.length;
        processedAccounts++;
      } else {
        Logger.log(`No transactions found for account ${accountId}`);
        processedAccounts++;
      }
    } catch (error) {
      Logger.log(`Error processing account ${accountId}: ${error.message}`);
      processedAccounts++;
    }
  });

  const resultMessage = `Processed ${processedAccounts} accounts. ` +
    `Loaded ${totalTransactions} transactions. ` +
    `${rateLimitedAccounts} accounts rate limited.`;
  SpreadsheetApp.getActive().toast(resultMessage, "Load Transactions Complete", 10);
}

function getAccountDataFromSpreadsheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet): Array<{ accountId: string; sheetName: string }> {
  const sheet = spreadsheet.getSheetByName("GoCardlessRequisitions");
  if (!sheet) {
    Logger.log("GoCardlessRequisitions sheet not found");
    SpreadsheetApp.getUi().alert("GoCardlessRequisitions sheet not found. Please link and fetch accounts first.");
    return [];
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // Assuming the account IDs are in the 5th column (index 4) and the Sheet Names are in the 6th column (index 5)
  const accountData = values.slice(1)
    .map(row => ({ accountId: row[4], sheetName: row[5] }));

  // Check if any account ID doesn't have a sheet name
  const missingSheetName = accountData.find(data => data.accountId && !data.sheetName);
  if (missingSheetName) {
    SpreadsheetApp.getUi().alert(`Account ID ${missingSheetName.accountId} doesn't have a sheet name. Please provide sheet names for all accounts.`);
    return [];
  }

  const validAccountData = accountData.filter(data => data.accountId && data.sheetName);

  if (validAccountData.length === 0) {
    SpreadsheetApp.getUi().alert("No valid accounts found. Please link and fetch accounts first.");
    return [];
  }

  Logger.log(`Found ${validAccountData.length} valid accounts with both account IDs and sheet names in the spreadsheet`);
  return validAccountData;
}
