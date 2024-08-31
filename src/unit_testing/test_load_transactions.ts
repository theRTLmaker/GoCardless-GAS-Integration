
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

    // Generate test transactions
    const testTransactions = generateTestTransactions();

    // Call storeTransactions with test data
    storeTransactions(spreadsheet, testAccountId, testSheetName, testTransactions);

    // Log completion message
    Logger.log(`Test completed. Check the "${testSheetName}" sheet for results.`);
    }
// Mock data
function generateTestTransactions(): Transaction[] {
    // Keep one transaction constant
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

    // Generate two random transactions
    const randomTransactions = [generateRandomTransaction(), generateRandomTransaction()];

    return [constantTransaction, ...randomTransactions];
}

function generateRandomTransaction(): Transaction {
    const transactionId = `RAND${Math.random().toString(36).substring(2, 15)}`;
    const date = new Date(Date.now() - Math.floor(Math.random() * 30) * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
    const amount = (Math.random() * 1000 - 500).toFixed(2);

    return {
        transactionId: transactionId,
        bookingDate: date,
        valueDate: date,
        transactionAmount: {
            amount: amount,
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