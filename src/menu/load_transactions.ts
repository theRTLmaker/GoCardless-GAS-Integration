const FETCH_PERIOD_START_DAYS_AGO = 90;
const FETCH_PERIOD_LENGTH_DAYS = 90;

const MAX_EXECUTION_TIME = 1000 * 300; // 5 mins max

function loadTransactions() {
  documentLock(_loadTransactions);
}
function _loadTransactions() {
  // console.log('loadTransactions')

  const scriptStart = new Date();

  const ui = SpreadsheetApp.getUi();

  function needGracefulShutdown() {
    if (+new Date() - +scriptStart > MAX_EXECUTION_TIME) {
      ui.alert(
        "Script running out of time but there's still data to process. Please run the script again"
      );
      return true;
    }
    return false;
  }

  const accessToken = getAccessToken();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let requisitionsSheet = spreadsheet.getSheetByName(REQUISITIONS_SHEET_NAME);

  if (!requisitionsSheet) {
    ui.alert(
      "Please link an account before using this command",
      ui.ButtonSet.OK
    );
    return;
  }

  let accountsSheet = spreadsheet.getSheetByName(ACCOUNTS_SHEET_NAME);

  const requisitions = requisitionsSheet.getSheetValues(
    2,
    1,
    requisitionsSheet.getLastRow() - 1,
    3
  );

  let accounts = getAccounts(accountsSheet);

  const {
    v_ApprovedSymbol,
    v_PendingSymbol,
    v_BalanceAdjustment,
    v_StartingBalance,
    v_BreakSymbol,
  } = getReferenceValues(spreadsheet);

  const {
    trx_Dates,
    trx_Uuids,
    trx_Accounts,
    trx_Categories,
    trx_Statuses,
    trx_Inflows,
  } = getReferenceRanges(spreadsheet);

  function getFirstEmptyRow(range: GoogleAppsScript.Spreadsheet.Range) {
    return (
      (parseInt(
        Object.entries(range.getValues())
          .reverse()
          .find(([index, [cell]]) => cell !== "")?.[0]
      ) + 1 || 0) + range.getRow()
    );
    // return parseInt(Object.entries(range.getValues()).find(
    //   ([index, [cell]]) => cell === ''
    // )?.[0]) + range.getRow()
  }

  const prevRowNumber = getFirstEmptyRow(trx_Dates);
  const txnSheet = trx_Dates.getSheet();
  // txnSheet.activate()

  let txnRowNumber = prevRowNumber;

  const prevUuids = trx_Uuids.getValues().map(([cell]) => cell);

  let warned_about_name = true; // FIXME - false

  const prevAccountRows: string[] = trx_Accounts
    .getValues()
    .map(([cell]) => cell);
  const prevDateRows: Date[] = trx_Dates.getValues().map(([cell]) => cell);
  const prevStatusRows: string[] = trx_Statuses
    .getValues()
    .map(([cell]) => cell);

  for (const { id: accountId, name: accountName, institutionId } of accounts) {
    if (needGracefulShutdown()) return;
    let account, balances;
    try {
      ({ account } = goCardlessRequest<any>(
        "/api/v2/accounts/" + encodeURIComponent(accountId) + "/details/",
        {
          headers: {
            Authorization: "Bearer " + accessToken,
          },
        }
      ));
      ({ balances } = goCardlessRequest<{
        balances: {
          balanceAmount: { amount: string };
          referenceDate: string;
          balanceType: string;
        }[];
      }>("/api/v2/accounts/" + encodeURIComponent(accountId) + "/balances/", {
        headers: {
          Authorization: "Bearer " + accessToken,
        },
      }));
    } catch (error) {
      // Update account status
      updateAccount(accountsSheet, {
        id: accountId,
        status: "ERROR",
        message: error.detail || error.message,
      });
      continue;
    }

    let balance = balances.find(
      ({ balanceType }) => balanceType === "interimCleared"
    );
    if (!balance)
      balance = balances.find(
        ({ balanceType }) => balanceType === "interimBooked"
      );
    if (!balance)
      balance = balances.find(
        ({ balanceType }) => balanceType === "interimAvailable"
      );
    if (!balance)
      balance = balances.find(({ balanceType }) => balanceType === "expected");

    // console.log(balances, 'Selected: ', balance?.balanceType);

    updateAccount(accountsSheet, {
      id: accountId,
      lastBalance: balance?.balanceAmount.amount,
      lastBalanceDate: balance?.referenceDate,
    });

    if (!accountName) {
      // accountsSheet.activate();
      if (!warned_about_name) {
        warned_about_name = true;
        SpreadsheetApp.getUi().alert(
          "An account doesn't have a name and is skipped. Please make sure to give it a name in the Nordigen Accounts sheet then try fetching again.",
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
      continue;
    }

    const accountDates = prevDateRows
      .filter(
        (_, index) =>
          prevAccountRows[index] === accountName &&
          (prevStatusRows[index] === v_ApprovedSymbol ||
            prevStatusRows[index] === v_BreakSymbol)
      )
      .filter((v) => !!v);
    // .filter((v,idx,arr) => arr.indexOf(v)===idx); // Dedupe will only work if this is cast to string

    let max_historical_days = FETCH_PERIOD_START_DAYS_AGO;
    {
      console.log("account", account);
      // Check for earlist access date
      const institutionsSheet = spreadsheet.getSheetByName(
        INSTITUTIONS_SHEET_NAME
      );
      const institutionIdx = institutionsSheet
        .getRange(2, 1, institutionsSheet.getLastRow(), 1)
        .getValues()
        .map(([cell]) => cell)
        .indexOf(institutionId);

      if (institutionIdx) {
        max_historical_days =
          institutionsSheet.getRange(institutionIdx + 2, 3, 1, 1).getValue() ||
          FETCH_PERIOD_START_DAYS_AGO;
      }
    }

    // start & end date
    const today = formatDate(new Date());
    const startDate = accountDates.length
      ? formatDate(new Date(Math.max(...accountDates.map((d) => +d))))
      : formatDate(addDays(new Date(), -max_historical_days));
    const endDate = minDate(
      formatDate(addDays(parseDate(startDate), FETCH_PERIOD_LENGTH_DAYS)),
      today
    );

    const { transactions } = goCardlessRequest<any>(
      "/api/v2/accounts/" +
        encodeURIComponent(accountId) +
        "/transactions/?date_from=" +
        encodeURIComponent(startDate) +
        "&date_to=" +
        encodeURIComponent(endDate),
      {
        headers: {
          Authorization: "Bearer " + accessToken,
        },
      }
    );

    for (const transaction of transactions.booked.slice().reverse()) {
      // if a transaction is seen before, update it
      let currentTxnRow;
      let prevIndex = transaction.transactionId
        ? prevUuids.indexOf(transaction.transactionId)
        : -1;
      if (prevIndex >= 0) {
        currentTxnRow = prevIndex + trx_Uuids.getRow();
        // console.log('txn exists')
      } else {
        currentTxnRow = txnRowNumber++;
        // console.log('new txn')
      }

      if (!transaction.transactionId) {
        console.log("transaction with no id", transaction);
      }

      // console.log('Inserting balance at row ', txnRowNumber)
      // console.log({transaction})

      updateTransaction(spreadsheet, currentTxnRow, {
        date: transaction.valueDate || transaction.bookingDate,
        amount: transaction.transactionAmount.amount,
        account: accountName,
        status: v_ApprovedSymbol as string,
        memo: [
          transaction.transactionAmount.amount > 0
            ? transaction.debtorName
            : transaction.creditorName,
          transaction.remittanceInformationUnstructured,
        ]
          .filter(Boolean)
          .join(" – "),
        uuid: transaction.transactionId,
      });
    }

    if (needGracefulShutdown()) return;

    if (balance) {
      // console.log('Inserting balance at row ', txnRowNumber)
      // FIXME - starting balance and balance adjustments

      if (
        endDate === today &&
        (!balance.referenceDate || balance.referenceDate === today)
      ) {
        // Check for an already existing starting balance, create one or do a balance adjustment
        const inflowRows: number[] = trx_Inflows
          .getValues()
          .map(([cell]) => cell);
        const accountRows: string[] = trx_Accounts
          .getValues()
          .map(([cell]) => cell);
        const dateRows: Date[] = trx_Dates.getValues().map(([cell]) => cell);
        const categoryRows: string[] = trx_Categories
          .getValues()
          .map(([cell]) => cell);

        const expectedBalance = parseFloat(balance.balanceAmount.amount);
        let startingBalance = inflowRows.find(
          (_, idx) =>
            accountRows[idx] === accountName &&
            categoryRows[idx] === v_StartingBalance
        );

        let currentBalance;
        {
          const calculationsSheet = spreadsheet.getSheetByName("Calculations");
          const calculationsAccounts = calculationsSheet
            .getRange("A7:A37")
            .getValues()
            .map(([cell]) => cell);
          const calculationsBalances = calculationsSheet
            .getRange("B7:B37")
            .getValues()
            .map(([cell]) => cell);

          let accountIdx = calculationsAccounts.indexOf(accountName);

          currentBalance = parseFloat(calculationsBalances[accountIdx]);
        }

        if (
          !Number.isNaN(currentBalance) &&
          expectedBalance !== currentBalance
        ) {
          if (startingBalance != null) {
            // Account already has a starting balance, make an adjustment
            updateTransaction(spreadsheet, txnRowNumber++, {
              date: balance.referenceDate || today,
              category: v_BalanceAdjustment as string,
              amount: expectedBalance - currentBalance,
              account: accountName,
              status: v_ApprovedSymbol as string,
            });
          } else {
            // Create a starting balance for the account
            let firstTransactionDate =
              dateRows
                .filter((_, index) => accountRows[index] === accountName)
                .map(formatDate)
                .sort()[0] || today;

            updateTransaction(spreadsheet, txnRowNumber++, {
              date: firstTransactionDate,
              category: v_StartingBalance as string,
              amount: expectedBalance - currentBalance,
              account: accountName,
              status: v_ApprovedSymbol as string,
            });
          }
        }
        console.log({
          currentBalance,
          startingBalance,
          newBalance: balance.balanceAmount.amount,
        });
      }
    }

    {
      // Check for most recent transactions, or insert a break if they're too old
      const accountRows = trx_Accounts.getValues().map(([cell]) => cell);
      const dateRows = trx_Dates.getValues().map(([cell]) => cell);

      let lastTransactionDate = dateRows
        .filter((_, index) => accountRows[index] === accountName)
        .map(formatDate)
        .sort()
        .reverse()[0];
      console.log({ lastTransactionDate, startDate, endDate });

      if (
        !lastTransactionDate ||
        (lastTransactionDate === startDate && today !== endDate)
      ) {
        // Insert a break, to avoid requesting the same period when there are no transactions
        updateTransaction(spreadsheet, txnRowNumber++, {
          date: endDate,
          account: accountName,
          status: v_BreakSymbol as string,
          memo: "Sync marker. Please don't remove",
        });
      }
    }

    updateAccount(accountsSheet, {
      id: accountId,
      lastFetched: endDate,
    });
  }

  const newRowNumber = getFirstEmptyRow(trx_Dates);

  if (prevRowNumber !== newRowNumber) {
    txnSheet
      .getRange(prevRowNumber, 2, newRowNumber - prevRowNumber, 7)
      .activate();
  }
}

function updateTransaction(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  rowNumber: number,
  {
    date,
    category,
    amount,
    account,
    status,
    memo,
    uuid,
  }: Partial<{
    date: string;
    category: string;
    amount: number;
    account: string;
    status: string;
    memo: string;
    uuid: string;
  }>
) {
  const {
    trx_Dates,
    trx_Categories,
    trx_Inflows,
    trx_Outflows,
    trx_Accounts,
    trx_Statuses,
    trx_Memos,
    trx_Uuids,
  } = getReferenceRanges(spreadsheet);

  const txnSheet = trx_Dates.getSheet();

  // some values should not be overriden unless specified (memo and category)

  txnSheet.getRange(rowNumber, trx_Dates.getColumn(), 1, 1).setValue(date);
  if (category)
    txnSheet
      .getRange(rowNumber, trx_Categories.getColumn(), 1, 1)
      .setValue(category);
  txnSheet
    .getRange(rowNumber, trx_Inflows.getColumn(), 1, 1)
    .setValue(amount >= 0 ? amount : "");
  txnSheet
    .getRange(rowNumber, trx_Outflows.getColumn(), 1, 1)
    .setValue(amount < 0 ? Math.abs(amount) : "");
  txnSheet
    .getRange(rowNumber, trx_Accounts.getColumn(), 1, 1)
    .setValue(account);
  txnSheet.getRange(rowNumber, trx_Statuses.getColumn(), 1, 1).setValue(status);
  if (memo)
    txnSheet.getRange(rowNumber, trx_Memos.getColumn(), 1, 1).setValue(memo);
  txnSheet.getRange(rowNumber, trx_Uuids.getColumn(), 1, 1).setValue(uuid);
}