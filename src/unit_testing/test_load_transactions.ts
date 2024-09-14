// Test wrapper function
function testStoreTransactions() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const testAccountId = "TEST_ACCOUNT_123";
    const testSheetName = "TestTransactions";
    const testBankName = "TestBank";

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
    storeTransactions(spreadsheet, testAccountId, testSheetName, testTransactions, testBankName);

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
        remittanceInformationUnstructuredArray: ["Constant transaction"],
        bankTransactionCode: "PMNT",
        creditorName: "John Doe", // Changed from debtorName to creditorName since amount is negative
        debtorAccount: {
            iban: "DE89370400440532013000"
        },
        isPending: false
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

    const isPositiveAmount = amount >= 0;
    const randomName = `${isPositiveAmount ? 'Debtor' : 'Creditor'} ${Math.random().toString(36).substring(2, 8)}`;

    return {
        transactionId: transactionId,
        bookingDate: date,
        valueDate: date,
        transactionAmount: {
            amount: amount.toFixed(2),
            currency: "EUR"
        },
        remittanceInformationUnstructuredArray: [`Random transaction ${transactionId}`],
        bankTransactionCode: "PMNT",
        ...(isPositiveAmount
            ? { debtorName: randomName }
            : { creditorName: randomName }
        ),
        debtorAccount: {
            iban: `DE${Math.floor(Math.random() * 1000000000000000000)}`
        },
        isPending: Math.random() < 0.5
    };
}

// Function to run the test
function runTransactionTest() {
    testStoreTransactions();
}