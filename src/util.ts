const CONFIG_SHEET_NAME = "GoCardlessData";
const SECRET_ID = "Secret ID";
const SECRET_KEY = "Secret Key";

const INSTITUTIONS_SHEET_NAME = "GoCardlessInstitutions";
const REQUISITIONS_SHEET_NAME = "GoCardlessRequisitions";
const ACCOUNTS_SHEET_NAME = "GoCardlessAccounts";


let config: [string, string][];

export function goCardlessRequest<T extends {}>(
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

export function getAccessToken() {
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
