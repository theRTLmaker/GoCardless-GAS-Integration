function linkAccount() {
  scriptLock(_linkAccount);
}

function _linkAccount() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Prompt for country code
  const countryResult = ui.prompt(
    "Enter country code",
    "e.g. GB for United Kingdom",
    ui.ButtonSet.OK_CANCEL
  );

  if (countryResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const countryCode = countryResult.getResponseText().toUpperCase();

  // Fetch institutions for the given country
  const accessToken = getAccessToken();
  const institutions = fetchInstitutions(accessToken, countryCode);

  if (institutions.length === 0) {
    ui.alert("No institutions found for the given country code.");
    return;
  }

  // Show institution selection prompt
  const selectedInstitution = showInstitutionSelectionPrompt(ui, institutions);

  if (!selectedInstitution) {
    return;
  }

  // Create agreement and requisition
  createAgreementAndRequisition(spreadsheet, accessToken, selectedInstitution);
}

function fetchInstitutions(accessToken: string, countryCode: string) {
  const url = `/api/v2/institutions/?country=${countryCode}`;
  return goCardlessRequest<Array<{ id: string; name: string }>>(url, {
    method: "get",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });
}

function showInstitutionSelectionPrompt(ui: GoogleAppsScript.Base.Ui, institutions: Array<{ id: string; name: string }>) {
  const options = institutions.map(inst => inst.name);
  const result = ui.prompt(
    "Select an Institution",
    "Enter the number of the institution you want to select:\n" +
    options.map((name, index) => `${index + 1}. ${name}`).join("\n"),
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const selectedIndex = parseInt(result.getResponseText()) - 1;
    if (selectedIndex >= 0 && selectedIndex < institutions.length) {
      return institutions[selectedIndex];
    }
  }
  return null;
}

function createAgreementAndRequisition(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, accessToken: string, institution: { id: string; name: string }) {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  const existingRequisitionId = scriptProperties.getProperty('LAST_REQUISITION_ID');

  if (existingRequisitionId) {
    const response = ui.alert(
      'Existing Account Link',
      'There is already an existing account link. Do you want to create a new one?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      ui.alert('Operation cancelled. You can use the existing link to fetch accounts.');
      return;
    }
  }

  const agreementData = goCardlessRequest<{ id: string }>("/api/v2/agreements/enduser/", {
    method: "post",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    payload: JSON.stringify({
      institution_id: institution.id,
      max_historical_days: 90,
      access_valid_for_days: 90,
      access_scope: ["balances", "details", "transactions"],
    }),
  });

  if (!agreementData || !agreementData.id) {
    throw new Error("Failed to create agreement");
  }

  Logger.log(`agreementData: ${JSON.stringify(agreementData)}`);

  const requisitionData = goCardlessRequest<{
    id: string;
    status: string;
    link: string;
  }>("/api/v2/requisitions/", {
    method: "post",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    payload: JSON.stringify({
      institution_id: institution.id,
      redirect: spreadsheet.getUrl(),
      agreement: agreementData.id,
    }),
  });

  Logger.log(`requisitionData: ${JSON.stringify(requisitionData)}`);

  // Store new requisition ID in script properties
  scriptProperties.setProperty('LAST_REQUISITION_ID', requisitionData.id);

  // Show authentication link to user
  const htmlOutput = HtmlService.createHtmlOutput(
    `<p>Go to <a href="${requisitionData.link}" target="_blank">this link</a> to authenticate your account.</p>` +
    `<p>Once done, close this dialog and use the "Fetch Accounts" option from the menu to load your account data.</p>`
  )
    .setWidth(450)
    .setHeight(250);

  ui.showModalDialog(htmlOutput, "Authenticate with your bank");
}