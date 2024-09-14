// // Compiled using gocardless-gas-integration 1.0.0 (TypeScript 4.9.5)
// function onInstall() {
//     onOpen();
//     // Perform additional setup as needed.
// }
// function onOpen() {
//   var ui = SpreadsheetApp.getUi();
//   SpreadsheetApp.getUi()
//       .createAddonMenu('GoCardless')
//       .addItem("Initialise", "initialise")
//       .addItem("Load institutions", "loadInstitutions")
//       .addItem("Link an account", "linkAccount")
//       .addItem("Load accounts", "loadAccounts")
//       .addItem("Load transactions", "loadTransactions")
//       .addSeparator()
//       .addSubMenu(SpreadsheetApp.getUi()
//       .createMenu("Utils")
//       .addItem("Clear empty transaction rows", "clearEmptyTransactionRows"))
//       .addToUi();
// }

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('GoCardless')
      .addItem('Link Account', 'linkAccount')
      .addItem('Fetch Accounts', 'fetchAccounts')
      .addItem('Load Transactions', 'loadTransactions')
      .addItem('Configure Transaction Columns', 'showColumnMappingDialog')
      .addItem('Sort Transactions and Update Balances', 'sortAndUpdateBalance')
      .addToUi();
  }