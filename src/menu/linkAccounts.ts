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

  // Show bank selection dialog
  showBankSelectionDialog(institutions);
}

function showBankSelectionDialog(institutions: Array<{ id: string; name: string; logo?: string }>) {
  try {
    const htmlTemplate = HtmlService.createTemplateFromFile('src/html/BankSelectionDialog');

    Logger.log(`Passing ${institutions.length} institutions to the HTML template`);

    // Pass the institutions to the template
    htmlTemplate.institutions = institutions;

    const html = htmlTemplate.evaluate()
      .setWidth(450)
      .setHeight(600);

    SpreadsheetApp.getUi().showModalDialog(html, 'Select a Bank');
  } catch (error) {
    Logger.log(`Error in showBankSelectionDialog: ${error.message}`);
    SpreadsheetApp.getUi().alert(`An error occurred while displaying the bank selection dialog: ${error.message}`);
  }
}

function selectBankAndContinue(bankId: string, bankName: string) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const accessToken = getAccessToken();
  createAgreementAndRequisition(spreadsheet, accessToken, { id: bankId, name: bankName });
}

function fetchInstitutions(accessToken: string, countryCode: string) {
  const url = `/api/v2/institutions/?country=${countryCode}`;
  return goCardlessRequest<Array<{ id: string; name: string; logo?: string }>>(url, {
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
    "Enter the number of the institution you want to select, or type 'sandbox' for a sandbox bank:\n" +
    options.map((name, index) => `${index + 1}. ${name}`).join("\n"),
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const input = result.getResponseText().trim().toLowerCase();
    if (input === 'sandbox') {
      return { id: 'SANDBOXFINANCE_SFIN0000', name: 'sandbox_bank' };
    }
    const selectedIndex = parseInt(input) - 1;
    if (selectedIndex >= 0 && selectedIndex < institutions.length) {
      return institutions[selectedIndex];
    } else {
      ui.alert('Invalid selection. Please enter a valid number or "sandbox".');
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
  const htmlTemplate = HtmlService.createTemplateFromFile('src/html/AuthenticationLinkDialog');
  htmlTemplate.authLink = requisitionData.link;

  const htmlOutput = htmlTemplate.evaluate()
    .setWidth(450)
    .setHeight(370);

  ui.showModalDialog(htmlOutput, "Authenticate with your bank");
}