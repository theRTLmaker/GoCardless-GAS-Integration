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
    SpreadsheetApp.getActive().toast("No accounts found or missing information. Please link and fetch accounts first, and provide sheet names and custom names for all accounts.", "Load Transactions");
    return;
  }

  const accessToken = getAccessToken();
  let totalTransactions = 0;
  let processedAccounts = 0;
  let rateLimitedAccounts = 0;

  accountData.forEach(({ accountId, sheetName, customName }) => {
    try {
      const transactions = fetchTransactionsForAccount(accessToken, accountId);
      if (transactions === null) {
        // Rate limit error occurred
        rateLimitedAccounts++;
      } else if (transactions && transactions.length > 0) {
        storeTransactions(spreadsheet, accountId, sheetName, transactions, customName);
        totalTransactions += transactions.length;
        processedAccounts++;
      } else {
        Logger.log(`No transactions found for account ${accountId} (${customName})`);
        processedAccounts++;
      }
    } catch (error) {
      Logger.log(`Error processing account ${accountId} (${customName}): ${error.message}`);
      processedAccounts++;
    }
  });

  const resultMessage = `Processed ${processedAccounts} accounts. ` +
    `Loaded ${totalTransactions} transactions. ` +
    `${rateLimitedAccounts} accounts rate limited.`;
  SpreadsheetApp.getActive().toast(resultMessage, "Load Transactions Complete", 10);
}

function getAccountDataFromSpreadsheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet): Array<{ accountId: string; sheetName: string; customName: string }> {
  const sheet = spreadsheet.getSheetByName(REQUISITIONS_SHEET_NAME);
  if (!sheet) {
    Logger.log(`${REQUISITIONS_SHEET_NAME} sheet not found`);
    SpreadsheetApp.getUi().alert(`${REQUISITIONS_SHEET_NAME} sheet not found. Please link and fetch accounts first.`);
    return [];
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // Assuming the account IDs are in the 5th column (index 4), the Sheet Names are in the 6th column (index 5),
  // and the Custom Account Names are in the 7th column (index 6)
  const accountData = values.slice(1)
    .map(row => ({ accountId: row[4], sheetName: row[5], customName: row[6] }));

  // Check if any account ID doesn't have a sheet name or custom name
  const missingInfo = accountData.find(data => data.accountId && (!data.sheetName || !data.customName));
  if (missingInfo) {
    SpreadsheetApp.getUi().alert(`Account ID ${missingInfo.accountId} is missing a sheet name or custom name. Please provide both for all accounts.`);
    return [];
  }

  const validAccountData = accountData.filter(data => data.accountId && data.sheetName && data.customName);

  if (validAccountData.length === 0) {
    SpreadsheetApp.getUi().alert("No valid accounts found. Please link and fetch accounts first, and provide sheet names and custom names for all accounts.");
    return [];
  }

  Logger.log(`Found ${validAccountData.length} valid accounts with account IDs, sheet names, and custom names in the spreadsheet`);
  return validAccountData;
}
