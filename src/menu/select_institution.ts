/// <reference types="@types/google-apps-script" />

import { getAccessToken, showSelectionPrompt, goCardlessRequest } from '../util.ts';
import { scriptLock } from '../lock.ts';

function selectInstitution(callbackFunctionName: string) {
  scriptLock(() => _selectInstitution(callbackFunctionName));
}

function _selectInstitution(callbackFunctionName: string) {
  const ui = SpreadsheetApp.getUi();

  let result = ui.prompt(
    "Please enter the country code:",
    ui.ButtonSet.OK_CANCEL
  );

  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.CANCEL || button == ui.Button.CLOSE) {
    return;
  }

  const accessToken = getAccessToken();
  Logger.log(`accessToken: ${accessToken}`);
  var url = "/api/v2/institutions/?country=" + text;
  Logger.log(`url: ${url}`);
  var institutionList = goCardlessRequest(url, {
    method: "get",
    headers: {
      Authorization: "Bearer " + accessToken,
      "Content-Type": "application/json",
    },
  }) as { name: string; id: string }[];

  Logger.log(`institutionList: ${institutionList}`);

  var names = institutionList.map(item => item.name);

  showSelectionPrompt(names, (selection: string) => {
    const selectedInstitution = institutionList.find(item => item.name === selection);
    if (selectedInstitution) {
      Logger.log(`Selected institution: ${JSON.stringify(selectedInstitution)}`);
      this[callbackFunctionName](selectedInstitution.id, selectedInstitution.name);
    } else {
      Logger.log(`No matching institution found for: ${selection}`);
    }
  }, "Select an Institution");
}