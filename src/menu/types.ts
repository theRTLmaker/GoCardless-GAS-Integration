export interface Transaction {
  transactionId?: string;
  bookingDate?: string;
  valueDate: string;
  transactionAmount: {
    amount: string;
    currency: string;
  };
  remittanceInformationUnstructured: string;
  bankTransactionCode?: string;
  debtorName?: string;
  debtorAccount?: {
    iban: string;
  };
}

// Add any other shared types here