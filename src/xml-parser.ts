/**
 * Tally XML Parser
 * Parses TallyPrime XML API responses into structured data.
 */

import { parseStringPromise } from 'xml2js';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface ParsedLedger {
  name: string;
  closingBalance: number;
}

export interface ParsedCompany {
  name: string;
  guid: string;
  startingFrom: string;
  booksFrom: string;
  lastVoucherDate: string;
}

export interface ParsedStockItem {
  name: string;
  parent: string;
  closingBalance: number;
  closingValue: number;
}

/** Internal representation of an xml2js node (loose dict). */
type XmlNode = Record<string, unknown>;

// ---------------------------------------------------------------------------
// Low-level xml2js helpers (all unsafe access is isolated here)
// ---------------------------------------------------------------------------

/**
 * Parse an XML string and return the COLLECTION node under
 * ENVELOPE > BODY > DATA > COLLECTION, or null if absent.
 *
 * xml2js returns loosely-typed `any` objects; the unsafe access is
 * intentionally confined to this single function.
 */
async function extractCollection(xml: string): Promise<XmlNode | null> {
  // xml2js returns loosely-typed objects; unsafe access is confined here.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-assignment
  const root: any = await parseStringPromise(xml, { explicitArray: true, mergeAttrs: false });
  // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-member-access
  const envelope: XmlNode | undefined = root?.ENVELOPE;
  if (!envelope) return null;
  // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-member-access
  const body: XmlNode = Array.isArray(envelope.BODY) ? (envelope.BODY as XmlNode[])[0] : (envelope.BODY as XmlNode);
  if (!body) return null;
  // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-member-access
  const data: XmlNode = Array.isArray(body.DATA) ? (body.DATA as XmlNode[])[0] : (body.DATA as XmlNode);
  if (!data) return null;
  // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-member-access
  const collection: XmlNode = Array.isArray(data.COLLECTION)
    ? (data.COLLECTION as XmlNode[])[0]
    : (data.COLLECTION as XmlNode);
  return collection ?? null;
}

/**
 * Safely extract the first string value from an xml2js parsed node.
 * xml2js represents text content as an array: `['text']` or `{ _: 'text' }`.
 */
function firstString(node: unknown): string {
  if (node === undefined || node === null) return '';
  if (Array.isArray(node)) {
    const first: unknown = node[0];
    if (typeof first === 'string') return first;
    if (typeof first === 'object' && first !== null && '_' in first) {
      return String((first as Record<string, unknown>)._);
    }
    return String(first ?? '');
  }
  if (typeof node === 'object' && '_' in (node as Record<string, unknown>)) {
    return String((node as Record<string, unknown>)._);
  }
  return String(node);
}

/**
 * Extract a named attribute from an xml2js `$` object.
 */
function attr(node: XmlNode, key: string): string {
  const attrs = node.$ as Record<string, string> | undefined;
  return attrs?.[key] ?? '';
}

/**
 * Parse a Tally amount string into a JavaScript number.
 * Tally amounts may contain commas; credit balances are negative.
 */
function parseTallyAmount(raw: string | undefined): number {
  if (!raw || raw.trim() === '') return 0;
  const cleaned = String(raw).replace(/,/g, '').trim();
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

// ---------------------------------------------------------------------------
// Public Parsers
// ---------------------------------------------------------------------------

/**
 * Parse the raw XML string from TallyPrime into structured ledger entries.
 * Handles the canonical response shape:
 *   <ENVELOPE><BODY><DATA><COLLECTION>
 *     <LEDGER NAME="..."><CLOSINGBALANCE TYPE="Amount">nnn</CLOSINGBALANCE></LEDGER>
 *   </COLLECTION></DATA></BODY></ENVELOPE>
 */
export async function parseLedgerResponse(xml: string): Promise<ParsedLedger[]> {
  const collection = await extractCollection(xml);
  if (!collection) return [];

  const ledgerNodes = (collection['LEDGER'] as XmlNode[] | undefined) ?? [];

  return ledgerNodes.map((n) => ({
    name: attr(n, 'NAME') || firstString(n['NAME']),
    closingBalance: parseTallyAmount(firstString(n['CLOSINGBALANCE'])),
  }));
}

/**
 * Sum the absolute values of all ledger closing balances.
 * Tally represents income/liability accounts as negative, so we take the
 * absolute value when computing totals for Sales / Purchase summaries.
 */
export function sumLedgerBalances(ledgers: ParsedLedger[]): number {
  return ledgers.reduce((sum, l) => sum + Math.abs(l.closingBalance), 0);
}

/**
 * Parse company list response from TallyPrime.
 */
export async function parseCompanyListResponse(xml: string): Promise<ParsedCompany[]> {
  const collection = await extractCollection(xml);
  if (!collection) return [];

  const companyNodes = (collection['COMPANY'] as XmlNode[] | undefined) ?? [];

  return companyNodes.map((n) => ({
    name: attr(n, 'NAME') || firstString(n['NAME']),
    guid: firstString(n['GUID']),
    startingFrom: firstString(n['STARTINGFROM']),
    booksFrom: firstString(n['BOOKSFROM']),
    lastVoucherDate: firstString(n['LASTVOUCHERDATE']),
  }));
}

/**
 * Parse stock item list response from TallyPrime.
 */
export async function parseStockItemResponse(xml: string): Promise<ParsedStockItem[]> {
  const collection = await extractCollection(xml);
  if (!collection) return [];

  const itemNodes =
    (collection['STOCKITEM'] as XmlNode[] | undefined) ??
    (collection['STOCK-ITEM'] as XmlNode[] | undefined) ??
    [];

  return itemNodes.map((n) => ({
    name: attr(n, 'NAME') || firstString(n['NAME']),
    parent: firstString(n['PARENT']),
    closingBalance: parseTallyAmount(firstString(n['CLOSINGBALANCE'])),
    closingValue: parseTallyAmount(firstString(n['CLOSINGVALUE'])),
  }));
}

