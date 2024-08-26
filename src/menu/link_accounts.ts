import { scriptLock } from '../lock';
import { getAccessToken, goCardlessRequest, INSTITUTIONS_SHEET_NAME, REQUISITIONS_SHEET_NAME } from '../util';

function linkAccount() {
  scriptLock(_linkAccount);
}

function _linkAccount() {
  const ui = SpreadsheetApp.getUi();

  let result = ui.prompt(
    "Please enter the institution ID:",
    ui.ButtonSet.OK_CANCEL
  );

  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.CANCEL || button == ui.Button.CLOSE) {
    return;
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let requisitionsSheet = spreadsheet.getSheetByName(REQUISITIONS_SHEET_NAME);

  if (!requisitionsSheet) {
    let activeSheet = spreadsheet.getActiveSheet();
    requisitionsSheet = spreadsheet
      .insertSheet()
      .setName(REQUISITIONS_SHEET_NAME);
    spreadsheet.setActiveSheet(requisitionsSheet);
    spreadsheet.moveActiveSheet(spreadsheet.getNumSheets());
    requisitionsSheet.appendRow(["ID", "Status", "Institution ID"]);
    spreadsheet.setActiveSheet(activeSheet);
    requisitionsSheet.hideSheet();
  }

  const accessToken = getAccessToken();
  console.log(`accessToken: ${accessToken}`);

  let agreementData: { id: string } | null = null;
  {
    // Create an agreement for max access
    const institutionsSheet = spreadsheet.getSheetByName(
      INSTITUTIONS_SHEET_NAME
    )!;
    const institutionIdx = institutionsSheet
      .getRange(2, 1, institutionsSheet.getLastRow(), 1)
      .getValues()
      .map(([cell]) => cell)
      .indexOf(text);

    if (institutionIdx) {
      const max_historical_days = institutionsSheet
        .getRange(institutionIdx + 2, 3, 1, 1)
        .getValue();
      const access_valid_for_days = 90;

      console.log({ max_historical_days, institutionIdx, text });

      agreementData = goCardlessRequest<{ id: string }>("/api/v2/agreements/enduser/", {
        method: "post",
        headers: {
          Authorization: "Bearer " + accessToken,
          "Content-Type": "application/json",
        },
        payload: JSON.stringify({
          institution_id: text,
          max_historical_days,
          access_valid_for_days,
          access_scope: ["balances", "details", "transactions"],
        }),
      });


    }
    if (!agreementData) {
      throw new Error("Failed to create agreement");
    }
    console.log({ agreementData });


  const data = goCardlessRequest<{
    id: string;
    status: string;
  }>("/api/v2/requisitions/", {
    method: "post",
    headers: {
      Authorization: "Bearer " + accessToken,
      "Content-Type": "application/json",
    },
    payload: JSON.stringify({
      institution_id: text,
      redirect: spreadsheet.getUrl(),
      agreement: agreementData?.id,
    }),
  });

  console.log({ data });

  requisitionsSheet.appendRow([data.id, data.status, text]);

  const htmlOutput = Object.assign(
    HtmlService.createTemplate(
      'Go to <a target="_blank" href="<?= data.link ?>">this link</a> to authenticate your account.<br>Once you\'re done come back here and you will be able to load your account transactions.<br>If the link doesnt work you can copy it and paste it in your browser:<pre><a target="_blank" href="<?=data.link?>"><?=data.link?></a></pre>'
    ),
    { data }
  )
    .evaluate()
    .setWidth(450)
    .setHeight(250);
    ui.showModalDialog(htmlOutput, "Authenticate with your bank");
  }
}