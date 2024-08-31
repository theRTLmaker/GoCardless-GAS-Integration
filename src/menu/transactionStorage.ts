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

  const lastRow = sheet.getLastRow();
  const transactionIdColumn = columnMappings['transactionId'];
  const transactionIdIndex = columnLetterToIndex(transactionIdColumn) - 1;

  let existingTransactionIds: Set<string> = new Set();
  if (lastRow > 1) {
    const existingIds = sheet.getRange(2, transactionIdIndex + 1, lastRow - 1, 1).getValues().flat();
    existingTransactionIds = new Set(existingIds.filter(id => id !== ""));
  }

  const newTransactions = transactions.filter(transaction => !existingTransactionIds.has(transaction.transactionId));

  if (newTransactions.length > 0) {
    const isSignalColumnSelected = 'transactionSignal' in columnMappings;
    const isCustomAccountNameSelected = 'customAccountName' in columnMappings;

    const dataToAppend = newTransactions.map(transaction => {
      const row = new Array(maxColumnIndex).fill('');
      Object.entries(columnIndexes).forEach(([field, index]) => {
        let value: any;
        if (field === 'transactionSignal') {
          const amount = parseFloat(getNestedValue(transaction, 'transactionAmount.amount'));
          value = amount >= 0 ? '+' : '-';
        } else if (field === 'customAccountName') {
          value = customName;
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

  sheet.autoResizeColumns(1, maxColumnIndex);

  Logger.log(`Stored ${newTransactions.length} new transactions for account ${accountId} (${customName}) in sheet ${sheetName}`);
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
    { field: 'transactionSignal', description: 'Signal', tooltip: 'The sign (+ or -) of the transaction amount.' }, // New field
    { field: 'transactionAmount.currency', description: 'Currency', tooltip: 'The currency in which the transaction amount is denominated.' },
    { field: 'remittanceInformationUnstructured', description: 'Remittance Info', tooltip: 'Additional information about the transaction, such as a payment reference or note.' },
    { field: 'bankTransactionCode', description: 'Transaction Code', tooltip: 'A code used by the bank to categorize the type of transaction.' },
    { field: 'debtorName', description: 'Debtor Name', tooltip: 'The name of the person or entity making the payment (for incoming transactions).' },
    { field: 'debtorAccount.iban', description: 'Debtor IBAN', tooltip: 'The International Bank Account Number of the debtor\'s account.' },
    { field: 'customAccountName', description: 'Custom Account Name', tooltip: 'The custom name assigned to this account in the Requisitions sheet.' }
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