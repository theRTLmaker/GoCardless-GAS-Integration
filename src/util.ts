export { goCardlessRequest, getAccessToken, showSelectionPrompt };

export const CONFIG_SHEET_NAME = "GoCardlessData";
export const SECRET_ID = "Secret ID";
export const SECRET_KEY = "Secret Key";

export const INSTITUTIONS_SHEET_NAME = "GoCardlessInstitutions";
export const REQUISITIONS_SHEET_NAME = "GoCardlessRequisitions";
export const ACCOUNTS_SHEET_NAME = "GoCardlessAccounts";


let config: [string, string][];

function goCardlessRequest<T extends {}>(
  url: string,
  { headers, ...options }: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions
): T {
  const request = UrlFetchApp.fetch("https://bankaccountdata.gocardless.com" + url, {
    ...options,
    headers: {
      accept: "application/json",
      ...headers,
    },
    muteHttpExceptions: true,
  });
  let data;
  try {
    Logger.log(request.getContentText());
    data = JSON.parse(request.getContentText());
  } catch (error) {
    throw new Error(
      "Unexpected response from the server. Code " + request.getResponseCode()
    );
  }
  if (request.getResponseCode() >= 400) {
    throw Object.assign(
      new Error("Unexpected server error. Code " + request.getResponseCode()),
      { statusCode: request.getResponseCode() },
      data
    );
  }
  return data;
}

function getAccessToken() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = spreadsheet.getSheetByName("GoCardlessData");

  if (!configSheet) {
    throw new Error(
      "GoCardless integration has not been initialised. Please initialise first."
    );
  }

  var config = configSheet.getSheetValues(
    1,
    1,
    configSheet.getLastRow(),
    2
  ) as string[][] as typeof config;

  const secret_id = config.find(([key]) => key === SECRET_ID)?.[1];
  if (!secret_id)
    throw new Error("Missing secret_id token from data. Please re-initialise");

  Logger.log(`secret_id: ${secret_id}`);

  const secret_key = config.find(([key]) => key === SECRET_KEY)?.[1];
  if (!secret_key)
    throw new Error("Missing secret_key token from data. Please re-initialise");

  Logger.log(`secret_key: ${secret_key}`);

  const data = goCardlessRequest<{ access: string }>("/api/v2/token/new/", {
    method: "post",
    headers: {
      "Content-Type": "application/json",
    },
    payload: JSON.stringify({
      "secret_id": secret_id,
      "secret_key": secret_key }),
  });
  Logger.log(`data: ${JSON.stringify(data, null, 2)}`);
  const { access } = data;

  Logger.log(`Access Token: ${access}`);
  return access;
}

function getReferenceValues(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet
) {
  return [
    "v_Today",
    "v_ReportableCategorySymbol",
    "v_NonReportableCategorySymbol",
    "v_DebtAccountSymbol",
    "v_CategoryGroupSymbol",
    "v_ApprovedSymbol",
    "v_PendingSymbol",
    "v_BreakSymbol",
    "v_AccountTransfer",
    "v_BalanceAdjustment",
    "v_StartingBalance",
  ].reduce((acc, name) => {
    Object.defineProperty(acc, name, {
      get() {
        const range = spreadsheet.getRangeByName(name);
        const value = range ? range.getValue() : null;
        Object.defineProperty(acc, name, { value });
        return value;
      },
      configurable: true,
    });
    return acc;
  }, {} as Record<string, string | number | Date | null>);
}

function getReferenceRanges(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet
) {
  return [
    "trx_Dates",
    "trx_Outflows",
    "trx_Inflows",
    "trx_Categories",
    "trx_Accounts",
    "trx_Statuses",
    "trx_Memos",
    "trx_Uuids",
    "ntw_Dates",
    "ntw_Amounts",
    "ntw_Categories",
    "cts_Dates",
    "cts_Amounts",
    "cts_FromCategories",
    "cts_ToCategories",
    "cfg_Accounts",
    "cfg_Cards",
  ].reduce((acc, name) => {
    Object.defineProperty(acc, name, {
      get() {
        const value = spreadsheet.getRangeByName(name);
        Object.defineProperty(acc, name, { value });
        return value;
      },
      configurable: true,
    });
    return acc;
  }, {} as Record<string, GoogleAppsScript.Spreadsheet.Range>);
}

function showSelectionPrompt(values: string[], onSelect: (selection: string) => void, title = 'Select an Option') {
  const ui = SpreadsheetApp.getUi();
  let html = '<html><body>';
  html += '<form id="myForm">';
  html += '<label>Select an option:</label><br><br>';

  // Populate the form with radio buttons
  values.forEach((value, i) => {
    html += `<input type="radio" id="option${i}" name="selection" value="${value}">`;
    html += `<label for="option${i}">${value}</label><br>`;
  });

  html += '<br><input type="button" value="Okay" onclick="submitSelection();" />';
  html += '</form>';
  html += '<script>';
  html += 'function getSelectedValue() {';
  html += '  const radios = document.getElementsByName("selection");';
  html += '  for (let i = 0; i < radios.length; i++) {';
  html += '    if (radios[i].checked) {';
  html += '      return radios[i].value;';
  html += '    }';
  html += '  }';
  html += '  return null;';
  html += '}';
  html += 'function submitSelection() {';
  html += '  const selection = getSelectedValue();';
  html += '  if (selection) {';
  html += '    google.script.run.withSuccessHandler(() => google.script.host.close()).processSelection(selection);';
  html += '  } else {';
  html += '    alert("Please select an option.");';
  html += '  }';
  html += '}';
  html += '</script>';
  html += '</body></html>';

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(400);

  // Show the dialog
  ui.showModalDialog(htmlOutput, title);

  // Set up the callback function
  this.processSelection = (selection: string) => {
    onSelect(selection);
  };
}