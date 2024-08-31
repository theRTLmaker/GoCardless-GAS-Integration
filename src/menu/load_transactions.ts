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

interface Transaction {
  transactionId?: string;
  bookingDate?: string;
  valueDate: string;
  transactionAmount: {
    amount: string;
    currency: string;
  };
  remittanceInformationUnstructured: string;
  bankTransactionCode?: string;
  debtorName?: string;
  debtorAccount?: {
    iban: string;
  };
}

function fetchTransactionsForAccount(accessToken: string, accountId: string): Transaction[] | null {
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
  const dateFrom = thirtyDaysAgo.toISOString().split('T')[0]; // Format: YYYY-MM-DD

  const url = `/api/v2/accounts/${accountId}/transactions/?date_from=${dateFrom}`;
  Logger.log(`Fetching transactions for account ${accountId} from ${dateFrom}`);

  try {
    const response = goCardlessRequest<{
      transactions: {
        booked: Transaction[],
        pending: Transaction[]
      }
    } | {
      summary: string;
      detail: string;
      status_code: number
    }>(url, {
      method: "get",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });

    // Logger.log(`Raw response for account ${accountId}: ${JSON.stringify(response)}`);

    if ('status_code' in response && response.status_code === 429) {
      const errorMessage = `Rate limit exceeded for account ${accountId}: ${response.summary}. ${response.detail}`;
      Logger.log(errorMessage);
      SpreadsheetApp.getActive().toast(errorMessage, "Rate Limit Error", 10);
      return null; // Indicate rate limit error
    }

    if (!('transactions' in response) || !response.transactions) {
      Logger.log(`Invalid response for account ${accountId}: No transactions found in the response`);
      return [];
    }

    const bookedTransactions = response.transactions.booked || [];
    const pendingTransactions = response.transactions.pending || [];
    const allTransactions = [...bookedTransactions, ...pendingTransactions];

    Logger.log(`Fetched ${bookedTransactions.length} booked and ${pendingTransactions.length} pending transactions for account ${accountId}`);
    return allTransactions;
  } catch (error) {
    Logger.log(`Error fetching transactions for account ${accountId}: ${error.message}`);
    throw error; // Rethrow the error to be caught in _loadTransactions
  }
}

function storeTransactions(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, accountId: string, sheetName: string, transactions: Transaction[]) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }

    const columnMappings = getCustomColumnMappings();
    const headers = Object.values(columnMappings);

    // Add headers if the sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    }

    const transactionRows = transactions.map(t =>
      Object.keys(columnMappings).map(key => {
        const value = key.split('.').reduce((obj, k) => obj && obj[k], t);
        return value !== undefined ? value : '';
      })
    );

    sheet.getRange(sheet.getLastRow() + 1, 1, transactionRows.length, headers.length).setValues(transactionRows);
    sheet.autoResizeColumns(1, headers.length);

    Logger.log(`Stored ${transactions.length} transactions for account ${accountId} in sheet ${sheetName}`);
  }

function getCustomColumnMappings(): Record<string, string> {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);

  if (!sheet) {
    throw new Error(`Sheet "${CONFIG_SHEET_NAME}" not found. Please run the initialization first.`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < COLUMN_CONFIG_START_ROW) {
    return {}; // No column config data yet
  }

  const data = sheet.getRange(COLUMN_CONFIG_START_ROW, 1, lastRow - COLUMN_CONFIG_START_ROW + 1, 2).getValues();
  return Object.fromEntries(data);
}

function showColumnMappingDialog() {
  const currentMappings = getCustomColumnMappings();

  // Convert the mappings to a JSON string
  const mappingsJson = JSON.stringify(currentMappings);

  // Read the HTML file content
  let htmlContent = HtmlService.createHtmlOutputFromFile('src/html/ColumnMappingDialog').getContent();

  // Replace a placeholder in the HTML with the mappings JSON
  htmlContent = htmlContent.replace('{{SAVED_MAPPINGS}}', mappingsJson);

  // Create a new HtmlOutput with the modified content
  const html = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(450)
      .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'Configure Transaction Columns');
}

function saveColumnMappings(mappings: { [key: string]: string }) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = spreadsheet.getSheetByName("TransactionConfig");
  if (!configSheet) {
    configSheet = spreadsheet.insertSheet("TransactionConfig");
  }

  const data = Object.entries(mappings).map(([key, value]) => [key, value]);
  configSheet.clear();
  configSheet.getRange(1, 1, data.length, 2).setValues(data);
}

// This function would be called from your HTML dialog
function updateColumnMappings(mappings: Record<string, string>) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);

  if (!sheet) {
    throw new Error(`Sheet "${CONFIG_SHEET_NAME}" not found. Please run the initialization first.`);
  }

  // Clear existing column config (starting from the third row)
  const lastRow = Math.max(sheet.getLastRow(), COLUMN_CONFIG_START_ROW);
  if (lastRow >= COLUMN_CONFIG_START_ROW) {
    sheet.getRange(COLUMN_CONFIG_START_ROW, 1, lastRow - COLUMN_CONFIG_START_ROW + 1, 2).clear();
  }

  // Save new config
  const configData = Object.entries(mappings).map(([field, column]) => [field, column]);
  sheet.getRange(COLUMN_CONFIG_START_ROW, 1, configData.length, 2).setValues(configData);
}

function getTransactionFieldsWithDescriptions(): Array<{field: string, description: string, tooltip: string}> {
  return [
    { field: 'transactionId', description: 'Transaction ID', tooltip: 'A unique identifier for each transaction.' },
    { field: 'bookingDate', description: 'Booking Date', tooltip: 'The date when the transaction was officially recorded by the bank.' },
    { field: 'valueDate', description: 'Value Date', tooltip: 'The date when the funds were actually debited or credited to the account.' },
    { field: 'transactionAmount.amount', description: 'Amount', tooltip: 'The monetary value of the transaction.' },
    { field: 'transactionAmount.currency', description: 'Currency', tooltip: 'The currency in which the transaction amount is denominated.' },
    { field: 'remittanceInformationUnstructured', description: 'Remittance Info', tooltip: 'Additional information about the transaction, such as a payment reference or note.' },
    { field: 'bankTransactionCode', description: 'Transaction Code', tooltip: 'A code used by the bank to categorize the type of transaction.' },
    { field: 'debtorName', description: 'Debtor Name', tooltip: 'The name of the person or entity making the payment (for incoming transactions).' },
    { field: 'debtorAccount.iban', description: 'Debtor IBAN', tooltip: 'The International Bank Account Number of the debtor\'s account.' }
  ];
}
