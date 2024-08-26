import { CONFIG_SHEET_NAME, SECRET_ID, SECRET_KEY } from '../util';

import { scriptLock } from '../lock';

function initialise() {
  scriptLock(_initialise);
}

function _initialise() {
  console.log("Initialise");

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let configSheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);
  if (configSheet) {
    let result = ui.alert(
      "GoCardless config already exists. Do you want to override it?",
      ui.ButtonSet.YES_NO
    );
    switch (result) {
      case ui.Button.NO:
      case ui.Button.CLOSE:
        return;
      case ui.Button.YES:
        spreadsheet.deleteSheet(configSheet);
        break;
    }
  }

  let result = ui.prompt(
    "Please enter your GoCardless ID:",
    ui.ButtonSet.OK_CANCEL
  );

  // Process the user's response.
  var button = result.getSelectedButton();
  var secret_id = result.getResponseText();
  if (button == ui.Button.CANCEL || button == ui.Button.CLOSE) {
    return;
  }

  result = ui.prompt(
    "Please enter your GoCardless Key:",
    ui.ButtonSet.OK_CANCEL
  );

  // Process the user's response.
  button = result.getSelectedButton();
  var secret_key = result.getResponseText();
  if (button == ui.Button.CANCEL || button == ui.Button.CLOSE) {
    return;
  }

  let activeSheet = spreadsheet.getActiveSheet();
  Logger.log("Creating new config sheet")
  configSheet = spreadsheet.insertSheet().setName(CONFIG_SHEET_NAME);

  configSheet.appendRow([SECRET_ID, secret_id]);
  configSheet.appendRow([SECRET_KEY, secret_key]);

  spreadsheet.setActiveSheet(activeSheet);
  configSheet.hideSheet();
}