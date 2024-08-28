function fetchAccounts() {
    const ui = SpreadsheetApp.getUi();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const requisitionId = PropertiesService.getScriptProperties().getProperty('LAST_REQUISITION_ID');

    if (!requisitionId) {
      ui.alert("No recent account link found. Please link an account first.");
      return;
    }

    const accessToken = getAccessToken();
    const accountIds = fetchAccountIds(accessToken, requisitionId);

    if (accountIds.length === 0) {
      ui.alert("No accounts found or authentication not completed. Please try again later.");
      return;
    }

    // Store account data
    storeRequisitionAndAccountData(spreadsheet, { id: requisitionId, status: 'COMPLETED' }, accountIds);

    ui.alert(`Successfully fetched ${accountIds.length} accounts.`);
  }

  function fetchAccountIds(accessToken: string, requisitionId: string): string[] {
    const url = `/api/v2/requisitions/${requisitionId}/`;
    const requisitionDetails = goCardlessRequest<{
      id: string;
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
      return [];
    }

    switch (requisitionDetails.status) {
      case 'CR':
        Logger.log(`Requisition ${requisitionId} is still in CREATED status. User needs to complete authentication.`);
        return [];
      case 'LN':
        Logger.log(`Requisition ${requisitionId} is in LINKED status.`);
        return requisitionDetails.accounts || [];
      case 'RJ':
        Logger.log(`Requisition ${requisitionId} was REJECTED.`);
        return [];
      case 'SA':
        Logger.log(`Requisition ${requisitionId} is in SUSPENDED state.`);
        return requisitionDetails.accounts || [];
      default:
        Logger.log(`Unknown status ${requisitionDetails.status} for requisition ${requisitionId}`);
        return [];
    }
  }

  function storeRequisitionAndAccountData(
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    requisitionData: { id: string; status: string },
    accountIds: string[]
  ) {
    let requisitionsSheet = spreadsheet.getSheetByName("GoCardlessRequisitions");
    if (!requisitionsSheet) {
      requisitionsSheet = spreadsheet.insertSheet("GoCardlessRequisitions");
      requisitionsSheet.appendRow(["ID", "Status", "Institution ID", "Institution Name", "Accounts", "Sheet Name"]);
    }

    // Find the institution data from the existing rows
    const dataRange = requisitionsSheet.getDataRange();
    const values = dataRange.getValues();
    let institutionId = "";
    let institutionName = "";

    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === requisitionData.id) {
        institutionId = values[i][2];
        institutionName = values[i][3];
        break;
      }
    }

    // Add a row for each account
    accountIds.forEach(accountId => {
      requisitionsSheet.appendRow([
        requisitionData.id,
        requisitionData.status,
        institutionId,
        institutionName,
        accountId,
        ""  // Empty Sheet Name column
      ]);
    });

    Logger.log(`Stored requisition and account data for ${accountIds.length} accounts`);

    // Optionally, you can auto-resize columns to fit the content
    requisitionsSheet.autoResizeColumns(1, 6);
  }