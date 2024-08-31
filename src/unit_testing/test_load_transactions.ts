// Test wrapper function
function testStoreTransactions() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const testAccountId = "TEST_ACCOUNT_123";
    const testSheetName = "TestTransactions";

    // Create a test sheet if it doesn't exist
    let sheet = spreadsheet.getSheetByName(testSheetName);
    if (!sheet) {
        sheet = spreadsheet.insertSheet(testSheetName);
    }

    // Set up test column mappings
    const testColumnMappings = {
        'transactionId': 'A',
        'bookingDate': 'B',
        'transactionAmount.amount': 'C',
        'transactionSignal': 'D',
        'transactionAmount.currency': 'E',
        'remittanceInformationUnstructured': 'F'
    };
    PropertiesService.getScriptProperties().setProperty('COLUMN_MAPPINGS', JSON.stringify(testColumnMappings));

    // Generate test transactions
    const testTransactions = generateTestTransactions();

    // Call storeTransactions with test data
    storeTransactions(spreadsheet, testAccountId, testSheetName, testTransactions);

    // Log completion message
    Logger.log(`Test completed. Check the "${testSheetName}" sheet for results.`);
}

// Mock data
function generateTestTransactions(): Transaction[] {
    const constantTransaction: Transaction = {
        transactionId: "CONST123456789",
        bookingDate: "2023-05-01",
        valueDate: "2023-05-01",
        transactionAmount: {
            amount: "-50.00",
            currency: "EUR"
        },
        remittanceInformationUnstructured: "Constant transaction",
        bankTransactionCode: "PMNT",
        debtorName: "John Doe",
        debtorAccount: {
            iban: "DE89370400440532013000"
        }
    };

    const randomTransactions = [
        generateRandomTransaction(true),  // Positive amount
        generateRandomTransaction(false), // Negative amount
        generateRandomTransaction()       // Random sign
    ];

    return [constantTransaction, ...randomTransactions];
}

function generateRandomTransaction(isPositive?: boolean): Transaction {
    const transactionId = `RAND${Math.random().toString(36).substring(2, 15)}`;
    const date = new Date(Date.now() - Math.floor(Math.random() * 30) * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
    let amount: number;

    if (isPositive === undefined) {
        amount = (Math.random() * 1000 - 500);
    } else {
        amount = isPositive ? Math.random() * 500 : -Math.random() * 500;
    }

    return {
        transactionId: transactionId,
        bookingDate: date,
        valueDate: date,
        transactionAmount: {
            amount: amount.toFixed(2),
            currency: "EUR"
        },
        remittanceInformationUnstructured: `Random transaction ${transactionId}`,
        bankTransactionCode: "PMNT",
        debtorName: `Random Debtor ${transactionId.substring(0, 5)}`,
        debtorAccount: {
            iban: `DE${Math.floor(Math.random() * 1000000000000000000)}`
        }
    };
}

// Function to run the test
function runTransactionTest() {
    testStoreTransactions();
}