declare function selectInstitution(callbackFunctionName: string): void;
declare function scriptLock(fn: Function): void;
declare function getAccessToken(): string;
declare function showSelectionPrompt(options: string[], callback: (selection: string) => void, title: string): void;
declare function goCardlessRequest<T>(url: string, options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions): T;