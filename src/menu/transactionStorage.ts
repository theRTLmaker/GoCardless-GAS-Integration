export function storeTransactions(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, accountId: string, sheetName: string, transactions: Transaction[], customName: string) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  const isNewSheet = !sheet;
  if (isNewSheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  const columnMappings = getCustomColumnMappings();
  if (Object.keys(columnMappings).length === 0) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Column Mappings Not Set', 'Please configure column mappings before storing transactions.', ui.ButtonSet.OK);
    showColumnMappingDialog();
    return;
  }

  const fieldDescriptions = getTransactionFieldsWithDescriptions().reduce((acc, field) => {
    acc[field.field] = field.description;
    return acc;
  }, {} as Record<string, string>);

  const columnIndexes = Object.fromEntries(
    Object.entries(columnMappings).map(([field, column]) => [field, columnLetterToIndex(column)])
  );

  const maxColumnIndex = Math.max(...Object.values(columnIndexes));

  if (isNewSheet || sheet.getLastRow() === 0) {
    const headerRow = new Array(maxColumnIndex).fill('');
    Object.entries(columnMappings).forEach(([field, column]) => {
      headerRow[columnLetterToIndex(column) - 1] = fieldDescriptions[field] || field;
    });
    sheet.getRange(1, 1, 1, maxColumnIndex).setValues([headerRow]);
  }

  const bookingDateColumn = columnMappings['bookingDate'];
  const bookingDateIndex = columnLetterToIndex(bookingDateColumn) - 1;
  const transactionIdColumn = columnMappings['transactionId'];
  const transactionIdIndex = columnLetterToIndex(transactionIdColumn) - 1;

  // Find the oldest date in the new transactions
  const oldestNewDate = transactions.reduce((oldest, transaction) => {
    const bookingDate = new Date(transaction.bookingDate);
    return bookingDate < oldest ? bookingDate : oldest;
  }, new Date());

  // Find the row with the oldest date that is the same or newer than the oldest new transaction
  let startRow = 2;
  if (sheet.getLastRow() > 1) {
    const dateValues = sheet.getRange(2, bookingDateIndex + 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < dateValues.length; i++) {
      const sheetDate = new Date(dateValues[i][0]);
      if (sheetDate >= oldestNewDate) {
        startRow = i + 1; // +1 because we start from row 2 and i is 0-indexed
        break;
      }
    }
  }

  let existingTransactionIds: Set<string> = new Set();
  if (startRow <= sheet.getLastRow()) {
    const existingIds = sheet.getRange(startRow, transactionIdIndex + 1, sheet.getLastRow() - startRow + 1, 1).getValues().flat();
    existingTransactionIds = new Set(existingIds.filter(id => id !== ""));
  }

  const newTransactions = transactions.filter(transaction => !existingTransactionIds.has(transaction.transactionId));

  // Delete pending transactions
  deletePendingTransactions(sheet, columnMappings);

  if (newTransactions.length > 0) {
    const lastRow = sheet.getLastRow();
    const isSignalColumnSelected = 'transactionSignal' in columnMappings;
    const isCustomAccountNameSelected = 'customAccountName' in columnMappings;
    const isTransactionStatusSelected = 'transactionStatus' in columnMappings;

    const dataToAppend = newTransactions.map(transaction => {
      const row = new Array(maxColumnIndex).fill('');
      Object.entries(columnIndexes).forEach(([field, index]) => {
        let value: any;
        if (field === 'transactionSignal') {
          const amount = parseFloat(getNestedValue(transaction, 'transactionAmount.amount'));
          value = amount >= 0 ? '+' : '-';
        } else if (field === 'customAccountName') {
          value = customName;
        } else if (field === 'transactionStatus') {
          value = transaction.isPending ? 'p' : '';
        } else if (field === 'debtorName') {
          // Replace debtorName with Merchant
          const amount = parseFloat(getNestedValue(transaction, 'transactionAmount.amount'));
          value = amount >= 0 ? transaction.debtorName : transaction.creditorName;
        } else if (field === 'remittanceInformationUnstructured') {
          // Join remittanceInformationUnstructuredArray
          value = transaction.remittanceInformationUnstructuredArray?.join(' ') || '';
        } else {
          value = getNestedValue(transaction, field);
        }

        if (value !== undefined) {
          if (field === 'transactionAmount.amount') {
            const amount = parseFloat(value);
            if (isSignalColumnSelected) {
              value = Math.abs(amount).toString(); // Store absolute value if signal column is selected
            } else {
              value = amount.toString(); // Keep the original value (with sign) if signal column is not selected
            }
          }
          row[index - 1] = value;
        }
      });
      return row;
    });

    sheet.getRange(lastRow + 1, 1, dataToAppend.length, maxColumnIndex).setValues(dataToAppend);
  }

  // Sort transactions by Booking Date
  sortTransactionsByBookingDate(sheet, columnMappings);

  // Update running balance
  updateRunningBalance(sheet, columnMappings);

  sheet.autoResizeColumns(1, maxColumnIndex);

  Logger.log(`Stored ${newTransactions.length} new transactions for account ${accountId} (${customName}) in sheet ${sheetName}`);
}

function deletePendingTransactions(sheet: GoogleAppsScript.Spreadsheet.Sheet, columnMappings: Record<string, string>) {
  const transactionStatusColumn = columnMappings['transactionStatus'];
  if (!transactionStatusColumn) return;

  const transactionStatusIndex = columnLetterToIndex(transactionStatusColumn);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log("Sheet is empty or only contains the header.");
    return;
  }
  const range = sheet.getRange(2, transactionStatusIndex, lastRow - 1, 1);
  const values = range.getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] === 'p') {
      sheet.deleteRow(i + 2); // +2 to account for header row and 0-based index
    }
  }
}

export function getCustomColumnMappings(): Record<string, string> {
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

export function showColumnMappingDialog() {
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

export function saveColumnMappings(mappings: { [key: string]: string }) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = spreadsheet.getSheetByName("TransactionConfig");
  if (!configSheet) {
    configSheet = spreadsheet.insertSheet("TransactionConfig");
  }

  const data = Object.entries(mappings).map(([key, value]) => [key, value]);
  configSheet.clear();
  configSheet.getRange(1, 1, data.length, 2).setValues(data);
}

export function updateColumnMappings(mappings: Record<string, string>) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);

  if (!sheet) {
    throw new Error(`Sheet "${CONFIG_SHEET_NAME}" not found. Please run the initialization first.`);
  }

  // Check for duplicate columns
  const columns = Object.values(mappings);
  const uniqueColumns = new Set(columns);
  if (columns.length !== uniqueColumns.size) {
    throw new Error("Duplicate columns detected. Each field must have a unique column.");
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

export function getTransactionFieldsWithDescriptions(): Array<{field: string, description: string, tooltip: string}> {
  return [
    { field: 'transactionId', description: 'Transaction ID', tooltip: 'A unique identifier for each transaction.' },
    { field: 'bookingDate', description: 'Booking Date', tooltip: 'The date when the transaction was officially recorded by the bank.' },
    { field: 'valueDate', description: 'Value Date', tooltip: 'The date when the funds were actually debited or credited to the account.' },
    { field: 'transactionAmount.amount', description: 'Amount', tooltip: 'The monetary value of the transaction.' },
    { field: 'transactionSignal', description: 'Signal', tooltip: 'The sign (+ or -) of the transaction amount.' },
    { field: 'transactionAmount.currency', description: 'Currency', tooltip: 'The currency in which the transaction amount is denominated.' },
    { field: 'remittanceInformationUnstructuredArray', description: 'Additional Info', tooltip: 'Additional information about the transaction, such as a payment reference or note.' },
    { field: 'bankTransactionCode', description: 'Transaction Code', tooltip: 'A code used by the bank to categorize the type of transaction.' },
    { field: 'debtorName', description: 'Merchant', tooltip: 'The merchant name for outgoing transactions or the debtor name for incoming transactions.' },
    { field: 'debtorAccount.iban', description: 'Debtor IBAN', tooltip: 'The International Bank Account Number of the debtor\'s account.' },
    { field: 'customAccountName', description: 'Custom Account Name', tooltip: 'The custom name assigned to this account in the Requisitions sheet.' },
    { field: 'transactionStatus', description: 'Transaction Status', tooltip: 'Indicates if the transaction is pending ("p") or booked (blank).' }
  ];
}

function getNestedValue(obj: any, path: string): any {
  return path.split('.').reduce((current, key) => current && current[key] !== undefined ? current[key] : undefined, obj);
}

function columnLetterToIndex(column: string): number {
  let index = 0;
  for (let i = 0; i < column.length; i++) {
    index = index * 26 + column.charCodeAt(i) - 64;
  }
  return index;
}

function sortTransactionsByBookingDate(sheet: GoogleAppsScript.Spreadsheet.Sheet, columnMappings: Record<string, string>) {
  const bookingDateColumn = columnMappings['bookingDate'];
  if (!bookingDateColumn) return;

  const bookingDateIndex = columnLetterToIndex(bookingDateColumn);
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  range.sort({ column: bookingDateIndex, ascending: true });
}

function updateRunningBalance(sheet: GoogleAppsScript.Spreadsheet.Sheet, columnMappings: Record<string, string>) {
  const amountColumn = columnMappings['transactionAmount.amount'];
  const customAccountNameColumn = columnMappings['customAccountName'];
  const signalColumn = columnMappings['transactionSignal'];
  if (!amountColumn || !customAccountNameColumn) return;

  const amountIndex = columnLetterToIndex(amountColumn);
  const customAccountNameIndex = columnLetterToIndex(customAccountNameColumn);
  const signalIndex = signalColumn ? columnLetterToIndex(signalColumn) : null;
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(2, amountIndex, lastRow - 1, 1);
  const values = range.getValues();

  const customAccountNames = sheet.getRange(2, customAccountNameIndex, lastRow - 1, 1).getValues().flat();
  const signals = signalIndex ? sheet.getRange(2, signalIndex, lastRow - 1, 1).getValues().flat() : null;

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const customAccountColumns: { [key: string]: number } = {};

  customAccountNames.forEach(name => {
    if (!customAccountColumns[name]) {
      let index = headerRow.indexOf(name) + 1;
      if (index === 0) {
        index = headerRow.length + 1;
        sheet.getRange(1, index).setValue(name);
      }
      customAccountColumns[name] = index;
    }
  });

  // Delete all balances from the Account Columns
  Object.values(customAccountColumns).forEach(columnIndex => {
    sheet.getRange(2, columnIndex, lastRow - 1, 1).clearContent();
  });

  const runningBalances: { [key: string]: number } = {};

  for (let i = 0; i < values.length; i++) {
    let amount = parseFloat(values[i][0]);
    const accountName = customAccountNames[i];
    if (!runningBalances[accountName]) {
      runningBalances[accountName] = 0;
    }
    if (signals) {
      amount = signals[i] === '-' ? -Math.abs(amount) : Math.abs(amount);
    }
    runningBalances[accountName] += amount;
    const columnIndex = customAccountColumns[accountName];
    sheet.getRange(i + 2, columnIndex).setValue(runningBalances[accountName]);
  }
}