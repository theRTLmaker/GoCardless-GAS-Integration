function initialise() {
  scriptLock(_initialise);
}

function _initialise() {
  Logger.log("Initialise");

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  if (configExists(spreadsheet)) {
    const shouldOverride = promptOverrideConfig(ui);
    if (!shouldOverride) {
      return;
    }
  }

  const credentials = promptForCredentials(ui);
  if (!credentials) {
    return;
  }

  createConfigSheet(spreadsheet, credentials);
}

function configExists(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet): boolean {
  return !!spreadsheet.getSheetByName(CONFIG_SHEET_NAME);
}

function promptOverrideConfig(ui: GoogleAppsScript.Base.Ui): boolean {
  const result = ui.alert(
    "GoCardless config already exists. Do you want to override it?",
    ui.ButtonSet.YES_NO
  );
  return result === ui.Button.YES;
}

function promptForCredentials(ui: GoogleAppsScript.Base.Ui): { secretId: string; secretKey: string } | null {
  const secretIdResult = ui.prompt(
    "Please enter your GoCardless ID:",
    ui.ButtonSet.OK_CANCEL
  );

  if (secretIdResult.getSelectedButton() !== ui.Button.OK) {
    return null;
  }

  const secretKeyResult = ui.prompt(
    "Please enter your GoCardless Key:",
    ui.ButtonSet.OK_CANCEL
  );

  if (secretKeyResult.getSelectedButton() !== ui.Button.OK) {
    return null;
  }

  return {
    secretId: secretIdResult.getResponseText(),
    secretKey: secretKeyResult.getResponseText()
  };
}

function createConfigSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, credentials: { secretId: string; secretKey: string }) {
  const activeSheet = spreadsheet.getActiveSheet();
  Logger.log("Creating new config sheet");

  let configSheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);
  if (configSheet) {
    spreadsheet.deleteSheet(configSheet);
  }

  configSheet = spreadsheet.insertSheet().setName(CONFIG_SHEET_NAME);

  configSheet.appendRow([SECRET_ID, credentials.secretId]);
  configSheet.appendRow([SECRET_KEY, credentials.secretKey]);

  spreadsheet.setActiveSheet(activeSheet);
  configSheet.hideSheet();

  Logger.log("Config sheet created and populated");
}