export interface Transaction {
  transactionId: string;
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
  isPending: boolean; // New field
}

// Add any other shared types here