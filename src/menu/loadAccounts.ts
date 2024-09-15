function fetchAccounts() {
    const ui = SpreadsheetApp.getUi();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const requisitionId = PropertiesService.getScriptProperties().getProperty('LAST_REQUISITION_ID');

    if (!requisitionId) {
      ui.alert("No recent account link found. Please link an account first.");
      return;
    }

    const accessToken = getAccessToken();
    const { accounts, institutionId } = fetchAccountIds(accessToken, requisitionId);

    if (accounts.length === 0) {
      ui.alert("No accounts found or authentication not completed. Please try again later.");
      return;
    }

    // Store account data
    storeRequisitionAndAccountData(spreadsheet, { id: requisitionId, status: 'COMPLETED' }, accounts, institutionId);

    ui.alert(`Successfully fetched ${accounts.length} accounts.`);
    PropertiesService.getScriptProperties().deleteProperty('LAST_REQUISITION_ID');
  }

  function fetchAccountIds(accessToken: string, requisitionId: string): { accounts: string[], institutionId: string } {
    const url = `/api/v2/requisitions/${requisitionId}/`;
    const requisitionDetails = goCardlessRequest<{
      id: string;
      institution_id: string;
      status: string;
      accounts: string[];
    }>(url, {
      method: "get",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });

    Logger.log(`requisitionDetails: ${JSON.stringify(requisitionDetails)}`);

    if (!requisitionDetails) {
      Logger.log(`Failed to fetch requisition details for ${requisitionId}`);
      return { accounts: [], institutionId: '' };
    }

    const result = { accounts: [], institutionId: requisitionDetails.institution_id };

    switch (requisitionDetails.status) {
      case 'CR':
        Logger.log(`Requisition ${requisitionId} is still in CREATED status. User needs to complete authentication.`);
        break;
      case 'LN':
        Logger.log(`Requisition ${requisitionId} is in LINKED status.`);
        result.accounts = requisitionDetails.accounts || [];
        break;
      case 'RJ':
        Logger.log(`Requisition ${requisitionId} was REJECTED.`);
        break;
      case 'SA':
        Logger.log(`Requisition ${requisitionId} is in SUSPENDED state.`);
        result.accounts = requisitionDetails.accounts || [];
        break;
      default:
        Logger.log(`Unknown status ${requisitionDetails.status} for requisition ${requisitionId}`);
    }

    return result;
  }

  function storeRequisitionAndAccountData(
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    requisitionData: { id: string; status: string },
    accountIds: string[],
    institutionId: string
  ) {
    let requisitionsSheet = spreadsheet.getSheetByName(REQUISITIONS_SHEET_NAME);
    if (!requisitionsSheet) {
      requisitionsSheet = spreadsheet.insertSheet(REQUISITIONS_SHEET_NAME);
      requisitionsSheet.appendRow(["ID", "Status", "Institution ID", "Accounts", "Sheet Name", "Custom Account Name", "Credit Card"]);
    }

    // Add a row for each account
    accountIds.forEach(accountId => {
      requisitionsSheet.appendRow([
        requisitionData.id,
        requisitionData.status,
        institutionId,
        accountId,
        "",  // Empty Sheet Name column
        "",  // Empty Custom Account Name column
        ""   // Empty Credit Card column
      ]);
    });

    Logger.log(`Stored requisition and account data for ${accountIds.length} accounts`);

    // Optionally, you can auto-resize columns to fit the content
    requisitionsSheet.autoResizeColumns(1, 7);
  }