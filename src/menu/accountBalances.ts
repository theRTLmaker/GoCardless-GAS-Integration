export function fetchAccountBalance(accessToken: string, accountId: string): number | null {
  const url = `/api/v2/accounts/${accountId}/balances/`;

  try {
    const response = goCardlessRequest<{
      balances: Array<{
        balanceAmount: {
          amount: string;
        }
      }>
    }>(url, {
      method: "get",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });

    if (response.balances && response.balances.length > 0) {
      return parseFloat(response.balances[0].balanceAmount.amount);
    } else {
      Logger.log(`No balance information found for account ${accountId}`);
      return null;
    }
  } catch (error) {
    Logger.log(`Error fetching balance for account ${accountId}: ${error instanceof Error ? error.message : String(error)}`);
    return null;
  }
}