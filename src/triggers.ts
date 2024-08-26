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
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('GoCardless')
        .addItem("Initialise", "initialise")
        .addItem("Link an account", "linkAccount")
        .addToUi();
  }