/**
 * Tally XML Request Templates
 * Generates correctly formatted TDL (Tally Definition Language) XML requests
 * for TallyPrime's HTTP XML API.
 */

// ---------------------------------------------------------------------------
// Date Helpers
// ---------------------------------------------------------------------------

/**
 * Format a date string (DD-MM-YYYY or YYYY-MM-DD) to TallyPrime's
 * required format: YYYYMMDD
 */
export function formatTallyDate(dateStr: string): string {
  if (/^\d{8}$/.test(dateStr)) {
    return dateStr; // already YYYYMMDD
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    // YYYY-MM-DD → YYYYMMDD
    return dateStr.replace(/-/g, '');
  }
  if (/^\d{2}-\d{2}-\d{4}$/.test(dateStr)) {
    // DD-MM-YYYY → YYYYMMDD
    const [dd, mm, yyyy] = dateStr.split('-');
    return `${yyyy}${mm}${dd}`;
  }
  return dateStr;
}

/**
 * Return { fromDate, toDate } strings in YYYYMMDD format for a given month/year.
 */
export function getMonthDates(month: number, year: number): { fromDate: string; toDate: string } {
  const mm = String(month).padStart(2, '0');
  const lastDay = new Date(year, month, 0).getDate();
  const dd = String(lastDay).padStart(2, '0');
  return {
    fromDate: `${year}${mm}01`,
    toDate: `${year}${mm}${dd}`,
  };
}

/**
 * Return { fromDate, toDate } for the full Indian financial year that begins
 * in `startYear` (i.e. 01-Apr-YYYY to 31-Mar-(YYYY+1)).
 */
export function getFinancialYearDates(startYear: number): { fromDate: string; toDate: string } {
  const endYear = startYear + 1;
  return {
    fromDate: `${startYear}0401`,
    toDate: `${endYear}0331`,
  };
}

// ---------------------------------------------------------------------------
// XML Request Builders
// ---------------------------------------------------------------------------

/**
 * Build an XML envelope that requests a named TDL collection.
 */
function buildCollectionRequest(params: {
  company: string;
  collectionName: string;
  collectionType: string;
  childOf?: string;
  fetch: string;
  fromDate?: string;
  toDate?: string;
  belongsTo?: string;
}): string {
  const { company, collectionName, collectionType, childOf, fetch, fromDate, toDate, belongsTo } =
    params;

  const dateVars =
    fromDate && toDate
      ? `
        <SVFROMDATE>${fromDate}</SVFROMDATE>
        <SVTODATE>${toDate}</SVTODATE>`
      : '';

  const childOfTag = childOf ? `\n            <CHILDOF>${childOf}</CHILDOF>` : '';
  const belongsToTag = belongsTo ? `\n            <BELONGSTO>${belongsTo}</BELONGSTO>` : '';

  return `<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>${collectionName}</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>${company}</SVCURRENTCOMPANY>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>${dateVars}
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="${collectionName}">
            <TYPE>${collectionType}</TYPE>${belongsToTag}${childOfTag}
            <FETCH>${fetch}</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>`;
}

/** List all companies currently loaded in TallyPrime. */
export function getCompanyListRequest(): string {
  return `<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>ListOfCompanies</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="ListOfCompanies">
            <TYPE>Company</TYPE>
            <FETCH>NAME, STARTINGFROM, BOOKSFROM, LASTVOUCHERDATE, GUID</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>`;
}

/** Get all Sales Account ledgers (for totalling sales revenue). */
export function getSalesAccountsRequest(
  company: string,
  fromDate: string,
  toDate: string
): string {
  return buildCollectionRequest({
    company,
    collectionName: 'SalesAccounts',
    collectionType: 'Ledger',
    childOf: 'Sales Accounts',
    fetch: 'NAME, CLOSINGBALANCE',
    fromDate,
    toDate,
  });
}

/** Get all Purchase Account ledgers (for totalling purchase costs). */
export function getPurchaseAccountsRequest(
  company: string,
  fromDate: string,
  toDate: string
): string {
  return buildCollectionRequest({
    company,
    collectionName: 'PurchaseAccounts',
    collectionType: 'Ledger',
    childOf: 'Purchase Accounts',
    fetch: 'NAME, CLOSINGBALANCE',
    fromDate,
    toDate,
  });
}

/** Get Sundry Debtors ledgers (outstanding receivables). */
export function getReceivablesRequest(company: string): string {
  return buildCollectionRequest({
    company,
    collectionName: 'SundryDebtors',
    collectionType: 'Ledger',
    childOf: 'Sundry Debtors',
    fetch: 'NAME, CLOSINGBALANCE',
  });
}

/** Get Sundry Creditors ledgers (outstanding payables). */
export function getPayablesRequest(company: string): string {
  return buildCollectionRequest({
    company,
    collectionName: 'SundryCreditors',
    collectionType: 'Ledger',
    childOf: 'Sundry Creditors',
    fetch: 'NAME, CLOSINGBALANCE',
  });
}

/** Get all stock items with closing quantities and values. */
export function getStockSummaryRequest(company: string): string {
  return buildCollectionRequest({
    company,
    collectionName: 'StockItems',
    collectionType: 'Stock Item',
    fetch: 'NAME, CLOSINGBALANCE, CLOSINGVALUE, PARENT',
  });
}

/** Get Trial Balance ledger data up to a given date. */
export function getTrialBalanceRequest(company: string, toDate: string): string {
  return buildCollectionRequest({
    company,
    collectionName: 'TrialBalance',
    collectionType: 'Ledger',
    fetch: 'NAME, PARENT, CLOSINGBALANCE',
    toDate,
  });
}

/** Get Profit & Loss account ledgers for a period. */
export function getProfitLossRequest(
  company: string,
  fromDate: string,
  toDate: string
): string {
  return buildCollectionRequest({
    company,
    collectionName: 'PLAccounts',
    collectionType: 'Ledger',
    childOf: 'Profit & Loss A/c',
    fetch: 'NAME, PARENT, CLOSINGBALANCE',
    fromDate,
    toDate,
  });
}
