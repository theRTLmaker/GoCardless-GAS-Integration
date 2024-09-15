export interface Transaction {
  transactionId: string;
  bookingDate: string;
  valueDate: string;
  transactionAmount: {
    amount: string;
    currency: string;
  };
  remittanceInformationUnstructuredArray?: string[];
  bankTransactionCode: string;
  debtorName?: string;
  creditorName?: string;
  additionalInformation?: string;
  debtorAccount: {
    iban: string;
  };
  isPending?: boolean;
  // ... any other fields you might have
}

// Add any other shared types here