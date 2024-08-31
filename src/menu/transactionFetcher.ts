function fetchTransactionsForAccount(accessToken: string, accountId: string): Transaction[] | null {
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
  const dateFrom = thirtyDaysAgo.toISOString().split('T')[0]; // Format: YYYY-MM-DD

  const url = `/api/v2/accounts/${accountId}/transactions/?date_from=${dateFrom}`;
  Logger.log(`Fetching transactions for account ${accountId} from ${dateFrom}`);

  try {
    const response = goCardlessRequest<{
      transactions: {
        booked: Transaction[],
        pending: Transaction[]
      }
    } | {
      summary: string;
      detail: string;
      status_code: number
    }>(url, {
      method: "get",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });

    Logger.log(`Raw response for account ${accountId}: ${JSON.stringify(response)}`);

    if ('transactions' in response && response.transactions) {
      const bookedTransactions = response.transactions.booked || [];
      const pendingTransactions = response.transactions.pending || [];
      const allTransactions = [...bookedTransactions, ...pendingTransactions];

      Logger.log(`Fetched ${bookedTransactions.length} booked and ${pendingTransactions.length} pending transactions for account ${accountId}`);
      return allTransactions;
    } else {
      Logger.log(`Invalid response for account ${accountId}: No transactions found in the response`);
      return [];
    }
  } catch (error) {
    if (error instanceof Error && 'statusCode' in error) {
      const statusCode = (error as any).statusCode;
      if (statusCode === 429) {
        const errorMessage = `Rate limit exceeded for account ${accountId}: ${error.message}`;
        Logger.log(errorMessage);
        SpreadsheetApp.getActive().toast(errorMessage, "Rate Limit Error", 10);
        return null; // Indicate rate limit error
      }
    }
    Logger.log(`Error fetching transactions for account ${accountId}: ${error instanceof Error ? error.message : String(error)}`);
    return []; // Return an empty array for other errors
  }
}