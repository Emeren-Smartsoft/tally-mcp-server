#!/usr/bin/env node
/**
 * Tally MCP Server
 * Model Context Protocol server for integrating TallyPrime accounting software
 * with Copilot Studio agents.
 *
 * Supports 25+ financial tools including monthly reports, annual reports,
 * operational reports, dashboard analytics, and utility functions.
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  Tool,
} from '@modelcontextprotocol/sdk/types.js';
import { randomUUID } from 'crypto';
import express, { NextFunction, Request, Response } from 'express';
import { z } from 'zod';
import { sendTallyRequest } from './tally-client.js';
import {
  getCompanyListRequest,
  getSalesAccountsRequest,
  getPurchaseAccountsRequest,
  getReceivablesRequest,
  getPayablesRequest,
  getStockSummaryRequest,
  getMonthDates,
} from './tally-requests.js';
import {
  parseLedgerResponse,
  parseCompanyListResponse,
  parseStockItemResponse,
  sumLedgerBalances,
} from './xml-parser.js';

// ---------------------------------------------------------------------------
// Type Definitions
// ---------------------------------------------------------------------------

interface TallyCompany {
  name: string;
  guid: string;
  startingFrom: string;
  booksFrom: string;
  lastVoucherDate: string;
}

interface LedgerEntry {
  ledgerName: string;
  groupName: string;
  openingBalance: number;
  debitAmount: number;
  creditAmount: number;
  closingBalance: number;
  currency: string;
}

interface VoucherEntry {
  date: string;
  voucherType: string;
  voucherNumber: string;
  partyName: string;
  amount: number;
  narration: string;
  ledgers: { name: string; amount: number; drCr: 'Dr' | 'Cr' }[];
}

interface TallyReport {
  companyName: string;
  reportDate: string;
  period: { from: string; to: string };
  data: Record<string, unknown>;
}

interface ConnectionState {
  method: 'odbc' | 'xml_api' | 'file' | 'demo';
  defaultCompany: string;
  connected: boolean;
}

// ---------------------------------------------------------------------------
// Zod Schemas for Input Validation
// ---------------------------------------------------------------------------

const DateRangeSchema = z.object({
  from_date: z
    .string()
    .regex(/^\d{2}-\d{2}-\d{4}$|^\d{4}-\d{2}-\d{2}$/, 'Date must be DD-MM-YYYY or YYYY-MM-DD')
    .describe('Start date (DD-MM-YYYY or YYYY-MM-DD)'),
  to_date: z
    .string()
    .regex(/^\d{2}-\d{2}-\d{4}$|^\d{4}-\d{2}-\d{2}$/, 'Date must be DD-MM-YYYY or YYYY-MM-DD')
    .describe('End date (DD-MM-YYYY or YYYY-MM-DD)'),
  company: z.string().optional().describe('Company name (uses default if not specified)'),
});

const MonthYearSchema = z.object({
  month: z.number().int().min(1).max(12).describe('Month number (1-12)'),
  year: z.number().int().min(2000).max(2100).describe('Four-digit year'),
  company: z.string().optional().describe('Company name (uses default if not specified)'),
});

const CompanySchema = z.object({
  company: z.string().min(1).describe('Company name'),
});

const LedgerQuerySchema = z.object({
  ledger_name: z.string().optional().describe('Specific ledger name'),
  group_name: z.string().optional().describe('Ledger group name'),
  from_date: z.string().optional().describe('Start date (DD-MM-YYYY or YYYY-MM-DD)'),
  to_date: z.string().optional().describe('End date (DD-MM-YYYY or YYYY-MM-DD)'),
  company: z.string().optional().describe('Company name (uses default if not specified)'),
});

const YearSchema = z.object({
  year: z.number().int().min(2000).max(2100).describe('Financial year (e.g., 2024 for FY 2024-25)'),
  company: z.string().optional().describe('Company name (uses default if not specified)'),
});

const AnalysisSchema = z.object({
  analysis_type: z
    .enum(['revenue', 'expense', 'profitability', 'liquidity', 'efficiency'])
    .describe('Type of financial analysis'),
  period: z.enum(['monthly', 'quarterly', 'annual']).describe('Analysis period'),
  year: z.number().int().min(2000).max(2100).describe('Financial year'),
  month: z.number().int().min(1).max(12).optional().describe('Month (required for monthly analysis)'),
  company: z.string().optional().describe('Company name (uses default if not specified)'),
});

// ---------------------------------------------------------------------------
// Connection State
// ---------------------------------------------------------------------------

const connectionState: ConnectionState = {
  method: 'xml_api',
  defaultCompany: process.env.TALLY_DEFAULT_COMPANY ?? 'Sample Company Ltd.',
  connected: false,
};

// ---------------------------------------------------------------------------
// Helper Utilities
// ---------------------------------------------------------------------------

/**
 * Format a number as Indian currency string (₹ X,XX,XXX.XX)
 */
function formatINR(amount: number): string {
  const absAmount = Math.abs(amount);
  const formatted = absAmount.toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  const prefix = amount < 0 ? '-₹' : '₹';
  return `${prefix}${formatted}`;
}

/**
 * Return month name for a given month number (1-12)
 */
function getMonthName(month: number): string {
  const months = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December',
  ];
  return months[month - 1] ?? 'Unknown';
}

/**
 * Format a date string to DD/MM/YYYY display format
 */
function formatDisplayDate(dateStr: string): string {
  // Handle YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    const [y, m, d] = dateStr.split('-');
    return `${d}/${m}/${y}`;
  }
  // Handle DD-MM-YYYY → DD/MM/YYYY
  return dateStr.replace(/-/g, '/');
}

/**
 * Derive the period start/end dates from a month/year
 */
function monthToDates(month: number, year: number): { from: string; to: string } {
  const lastDay = new Date(year, month, 0).getDate();
  const mm = String(month).padStart(2, '0');
  const dd = String(lastDay).padStart(2, '0');
  return {
    from: `01-${mm}-${year}`,
    to: `${dd}-${mm}-${year}`,
  };
}

/**
 * Resolve company name from optional input, falling back to default
 */
function resolveCompany(company?: string): string {
  return company ?? connectionState.defaultCompany;
}

// ---------------------------------------------------------------------------
// Demo Data Generators
// NOTE: In production these functions would execute ODBC/XML API queries
//       against TallyPrime.  The demo layer lets the server run and respond
//       without a live Tally installation.
// ---------------------------------------------------------------------------

function generateSalesData(month: number, year: number) {
  const base = 500000 + month * 12500;
  return {
    totalSales: base,
    totalTax: base * 0.18,
    netSales: base - base * 0.18,
    numberOfTransactions: 40 + month,
    topCustomers: [
      { name: 'ABC Corporation Pvt Ltd', amount: base * 0.22 },
      { name: 'XYZ Enterprises', amount: base * 0.17 },
      { name: 'Global Traders Ltd', amount: base * 0.13 },
      { name: 'Prime Solutions Inc', amount: base * 0.1 },
      { name: 'Sunrise Industries', amount: base * 0.08 },
    ],
    previousMonth: base * 0.92,
    growth: 8.7,
  };
}

function generatePurchaseData(month: number, year: number) {
  const base = 350000 + month * 8000;
  return {
    totalPurchases: base,
    totalTax: base * 0.18,
    netPurchases: base - base * 0.18,
    numberOfTransactions: 25 + month,
    topVendors: [
      { name: 'National Supplies Co', amount: base * 0.25 },
      { name: 'Metro Distributors', amount: base * 0.2 },
      { name: 'City Wholesalers', amount: base * 0.15 },
      { name: 'Trade Hub Pvt Ltd', amount: base * 0.12 },
      { name: 'Regional Suppliers', amount: base * 0.1 },
    ],
    previousMonth: base * 0.94,
    growth: 6.4,
  };
}

function generatePLData(month: number, year: number) {
  const sales = 500000 + month * 12500;
  const purchases = 350000 + month * 8000;
  const expenses = 85000 + month * 1500;
  const grossProfit = sales - purchases;
  const netProfit = grossProfit - expenses;
  return {
    totalIncome: sales,
    costOfGoodsSold: purchases,
    grossProfit,
    grossMargin: (grossProfit / sales) * 100,
    operatingExpenses: expenses,
    ebitda: netProfit + 15000,
    netProfit,
    netMargin: (netProfit / sales) * 100,
    incomeBreakdown: [
      { category: 'Sales Revenue', amount: sales * 0.9 },
      { category: 'Service Income', amount: sales * 0.1 },
    ],
    expenseBreakdown: [
      { category: 'Salaries & Wages', amount: expenses * 0.45 },
      { category: 'Rent & Utilities', amount: expenses * 0.2 },
      { category: 'Administrative', amount: expenses * 0.15 },
      { category: 'Marketing', amount: expenses * 0.12 },
      { category: 'Miscellaneous', amount: expenses * 0.08 },
    ],
  };
}

function generateCashFlowData(month: number, year: number) {
  const operating = 120000 + month * 3000;
  const investing = -25000 - month * 500;
  const financing = -15000;
  return {
    operatingActivities: {
      total: operating,
      items: [
        { description: 'Cash received from customers', amount: 580000 + month * 10000 },
        { description: 'Cash paid to suppliers', amount: -(350000 + month * 8000) },
        { description: 'Cash paid for expenses', amount: -(85000 + month * 1500) },
        { description: 'GST paid (net)', amount: -25000 },
      ],
    },
    investingActivities: {
      total: investing,
      items: [
        { description: 'Purchase of fixed assets', amount: investing * 0.8 },
        { description: 'Proceeds from asset sales', amount: Math.abs(investing) * 0.2 },
      ],
    },
    financingActivities: {
      total: financing,
      items: [
        { description: 'Loan repayments', amount: -10000 },
        { description: 'Dividend payments', amount: -5000 },
      ],
    },
    netCashFlow: operating + investing + financing,
    openingBalance: 350000,
    closingBalance: 350000 + operating + investing + financing,
  };
}

function generateGSTData(month: number, year: number) {
  const salesBase = 500000 + month * 12500;
  const purchaseBase = 350000 + month * 8000;
  return {
    outputGST: {
      igst: salesBase * 0.06,
      cgst: salesBase * 0.06,
      sgst: salesBase * 0.06,
      total: salesBase * 0.18,
    },
    inputGST: {
      igst: purchaseBase * 0.04,
      cgst: purchaseBase * 0.07,
      sgst: purchaseBase * 0.07,
      total: purchaseBase * 0.18,
    },
    netGSTLiability: salesBase * 0.18 - purchaseBase * 0.18,
    gstReturns: [
      { returnType: 'GSTR-1', dueDate: `11-${String(month + 1).padStart(2, '0')}-${year}`, status: 'Due' },
      { returnType: 'GSTR-3B', dueDate: `20-${String(month + 1).padStart(2, '0')}-${year}`, status: 'Due' },
    ],
    hsnSummary: [
      { hsn: '998311', description: 'IT Services', taxableValue: salesBase * 0.4, tax: salesBase * 0.4 * 0.18 },
      { hsn: '998314', description: 'Consulting', taxableValue: salesBase * 0.35, tax: salesBase * 0.35 * 0.18 },
      { hsn: '85423', description: 'Electronic Components', taxableValue: salesBase * 0.25, tax: salesBase * 0.25 * 0.18 },
    ],
  };
}

function generateLedgerSummaryData(month: number, year: number) {
  return {
    ledgers: [
      { name: 'Cash in Hand', group: 'Cash-in-Hand', opening: 45000, debit: 320000, credit: 280000, closing: 85000 },
      { name: 'HDFC Bank OD A/c', group: 'Bank OD Accounts', opening: -200000, debit: 500000, credit: 450000, closing: -150000 },
      { name: 'Accounts Receivable', group: 'Sundry Debtors', opening: 750000, debit: 620000, credit: 480000, closing: 890000 },
      { name: 'Accounts Payable', group: 'Sundry Creditors', opening: -420000, debit: 380000, credit: 350000, closing: -390000 },
      { name: 'Sales Account', group: 'Sales Accounts', opening: 0, debit: 0, credit: 500000 + month * 12500, closing: -(500000 + month * 12500) },
      { name: 'Purchase Account', group: 'Purchase Accounts', opening: 0, debit: 350000 + month * 8000, credit: 0, closing: 350000 + month * 8000 },
    ],
    totalDebits: 2170000,
    totalCredits: 2060000,
  };
}

function generateAnnualPLData(year: number) {
  const annualSales = 7200000;
  const annualPurchases = 4800000;
  const annualExpenses = 1200000;
  const grossProfit = annualSales - annualPurchases;
  const netProfit = grossProfit - annualExpenses;
  return {
    financialYear: `${year}-${String(year + 1).slice(-2)}`,
    totalIncome: annualSales,
    costOfGoodsSold: annualPurchases,
    grossProfit,
    grossMargin: (grossProfit / annualSales) * 100,
    operatingExpenses: annualExpenses,
    ebitda: netProfit + 180000,
    depreciation: 120000,
    ebit: netProfit + 60000,
    interestExpense: 60000,
    pbt: netProfit,
    taxProvision: netProfit * 0.25,
    pat: netProfit * 0.75,
    netMargin: ((netProfit * 0.75) / annualSales) * 100,
    quarterlyBreakdown: [
      { quarter: 'Q1 (Apr-Jun)', income: annualSales * 0.22, profit: netProfit * 0.18 },
      { quarter: 'Q2 (Jul-Sep)', income: annualSales * 0.24, profit: netProfit * 0.22 },
      { quarter: 'Q3 (Oct-Dec)', income: annualSales * 0.28, profit: netProfit * 0.27 },
      { quarter: 'Q4 (Jan-Mar)', income: annualSales * 0.26, profit: netProfit * 0.33 },
    ],
  };
}

function generateBalanceSheetData(year: number) {
  return {
    asOnDate: `31-03-${year + 1}`,
    assets: {
      fixedAssets: {
        grossBlock: 3500000,
        lessDepreciation: -850000,
        netBlock: 2650000,
        capitalWIP: 250000,
        total: 2900000,
      },
      currentAssets: {
        inventory: 650000,
        debtors: 1200000,
        cashAndBank: 450000,
        loansAndAdvances: 180000,
        otherCurrentAssets: 120000,
        total: 2600000,
      },
      totalAssets: 5500000,
    },
    liabilitiesAndEquity: {
      shareholdersFunds: {
        shareCapital: 1000000,
        reserves: 1800000,
        total: 2800000,
      },
      longTermLiabilities: {
        longTermBorrowings: 1200000,
        deferredTaxLiability: 150000,
        total: 1350000,
      },
      currentLiabilities: {
        shortTermBorrowings: 400000,
        creditors: 650000,
        otherCurrentLiabilities: 300000,
        total: 1350000,
      },
      totalLiabilitiesAndEquity: 5500000,
    },
    ratios: {
      currentRatio: 1.93,
      debtToEquity: 0.57,
      returnOnEquity: 18.5,
      returnOnAssets: 9.4,
    },
  };
}

function generateOutstandingData(type: 'receivable' | 'payable' | 'both') {
  const receivables = [
    { party: 'ABC Corporation Pvt Ltd', amount: 285000, daysOverdue: 45, ageing: '31-60 days' },
    { party: 'XYZ Enterprises', amount: 192000, daysOverdue: 18, ageing: '0-30 days' },
    { party: 'Global Traders Ltd', amount: 156000, daysOverdue: 72, ageing: '61-90 days' },
    { party: 'Prime Solutions Inc', amount: 98000, daysOverdue: 95, ageing: '>90 days' },
    { party: 'Sunrise Industries', amount: 67000, daysOverdue: 12, ageing: '0-30 days' },
  ];
  const payables = [
    { party: 'National Supplies Co', amount: 185000, daysOverdue: 22, ageing: '0-30 days' },
    { party: 'Metro Distributors', amount: 142000, daysOverdue: 35, ageing: '31-60 days' },
    { party: 'City Wholesalers', amount: 95000, daysOverdue: 8, ageing: '0-30 days' },
    { party: 'Trade Hub Pvt Ltd', amount: 78000, daysOverdue: 55, ageing: '31-60 days' },
  ];
  return { receivables, payables };
}

function generateStockData() {
  return {
    totalStockValue: 650000,
    totalItems: 48,
    groups: [
      {
        group: 'Finished Goods',
        items: [
          { name: 'Product A - Model X1', unit: 'Nos', quantity: 150, rate: 2500, value: 375000 },
          { name: 'Product B - Model Y2', unit: 'Nos', quantity: 80, rate: 1800, value: 144000 },
          { name: 'Product C - Model Z3', unit: 'Nos', quantity: 45, rate: 1200, value: 54000 },
        ],
      },
      {
        group: 'Raw Materials',
        items: [
          { name: 'Raw Material - Type A', unit: 'Kg', quantity: 500, rate: 85, value: 42500 },
          { name: 'Raw Material - Type B', unit: 'Ltrs', quantity: 200, rate: 175, value: 35000 },
        ],
      },
    ],
    lowStockItems: [
      { name: 'Product B - Model Y2', currentStock: 80, reorderLevel: 100 },
      { name: 'Raw Material - Type A', currentStock: 500, reorderLevel: 600 },
    ],
  };
}

// ---------------------------------------------------------------------------
// Report Formatters
// ---------------------------------------------------------------------------

async function formatMonthlySalesSummary(company: string, month: number, year: number): Promise<string> {
  const { from, to } = monthToDates(month, year);
  const { fromDate, toDate } = getMonthDates(month, year);

  try {
    const xml = await sendTallyRequest(getSalesAccountsRequest(company, fromDate, toDate));
    const ledgers = await parseLedgerResponse(xml);
    if (ledgers.length > 0) {
      const totalSales = sumLedgerBalances(ledgers);
      const lines: string[] = [
        `# Sales Summary - ${getMonthName(month)} ${year}`,
        `**Company:** ${company}`,
        `**Period:** ${formatDisplayDate(from)} to ${formatDisplayDate(to)}`,
        '',
        '## Key Metrics',
        `- **Total Sales:** ${formatINR(totalSales)}`,
        '',
        '## Sales Accounts',
        `| Ledger | Closing Balance |`,
        `|--------|----------------|`,
        ...ledgers.map((l) => `| ${l.name} | ${formatINR(Math.abs(l.closingBalance))} |`),
        `| **Total** | **${formatINR(totalSales)}** |`,
      ];
      return lines.join('\n');
    }
  } catch {
    // Fall through to demo data
  }

  // Demo fallback
  const data = generateSalesData(month, year);
  const lines: string[] = [
    `# Sales Summary - ${getMonthName(month)} ${year}`,
    `**Company:** ${company}`,
    `**Period:** ${formatDisplayDate(from)} to ${formatDisplayDate(to)}`,
    '',
    '## Key Metrics',
    `- **Total Sales (incl. Tax):** ${formatINR(data.totalSales)}`,
    `- **Net Sales (excl. Tax):** ${formatINR(data.netSales)}`,
    `- **Total Tax Collected:** ${formatINR(data.totalTax)}`,
    `- **Number of Transactions:** ${data.numberOfTransactions}`,
    `- **Previous Month Sales:** ${formatINR(data.previousMonth)}`,
    `- **Month-on-Month Growth:** ${data.growth > 0 ? '+' : ''}${data.growth.toFixed(1)}%`,
    '',
    '## Top Customers',
    ...data.topCustomers.map((c, i) => `${i + 1}. **${c.name}:** ${formatINR(c.amount)}`),
    '',
    '## Summary',
    `Sales for ${getMonthName(month)} ${year} ${data.growth > 0 ? 'grew' : 'declined'} by ${Math.abs(data.growth).toFixed(1)}% compared to the previous month.`,
    '',
    `_Note: TallyPrime not reachable — showing demo data._`,
  ];
  return lines.join('\n');
}

async function formatMonthlyPurchaseSummary(company: string, month: number, year: number): Promise<string> {
  const { from, to } = monthToDates(month, year);
  const { fromDate, toDate } = getMonthDates(month, year);

  try {
    const xml = await sendTallyRequest(getPurchaseAccountsRequest(company, fromDate, toDate));
    const ledgers = await parseLedgerResponse(xml);
    if (ledgers.length > 0) {
      const totalPurchases = sumLedgerBalances(ledgers);
      const lines: string[] = [
        `# Purchase Summary - ${getMonthName(month)} ${year}`,
        `**Company:** ${company}`,
        `**Period:** ${formatDisplayDate(from)} to ${formatDisplayDate(to)}`,
        '',
        '## Key Metrics',
        `- **Total Purchases:** ${formatINR(totalPurchases)}`,
        '',
        '## Purchase Accounts',
        `| Ledger | Closing Balance |`,
        `|--------|----------------|`,
        ...ledgers.map((l) => `| ${l.name} | ${formatINR(Math.abs(l.closingBalance))} |`),
        `| **Total** | **${formatINR(totalPurchases)}** |`,
      ];
      return lines.join('\n');
    }
  } catch {
    // Fall through to demo data
  }

  // Demo fallback
  const data = generatePurchaseData(month, year);
  const lines: string[] = [
    `# Purchase Summary - ${getMonthName(month)} ${year}`,
    `**Company:** ${company}`,
    `**Period:** ${formatDisplayDate(from)} to ${formatDisplayDate(to)}`,
    '',
    '## Key Metrics',
    `- **Total Purchases (incl. Tax):** ${formatINR(data.totalPurchases)}`,
    `- **Net Purchases (excl. Tax):** ${formatINR(data.netPurchases)}`,
    `- **Total Input Tax:** ${formatINR(data.totalTax)}`,
    `- **Number of Transactions:** ${data.numberOfTransactions}`,
    `- **Previous Month Purchases:** ${formatINR(data.previousMonth)}`,
    `- **Month-on-Month Growth:** ${data.growth > 0 ? '+' : ''}${data.growth.toFixed(1)}%`,
    '',
    '## Top Vendors',
    ...data.topVendors.map((v, i) => `${i + 1}. **${v.name}:** ${formatINR(v.amount)}`),
    '',
    `_Note: TallyPrime not reachable — showing demo data._`,
  ];
  return lines.join('\n');
}

function formatMonthlyPL(company: string, month: number, year: number): string {
  const data = generatePLData(month, year);
  const { from, to } = monthToDates(month, year);
  const lines: string[] = [
    `# Profit & Loss Statement - ${getMonthName(month)} ${year}`,
    `**Company:** ${company}`,
    `**Period:** ${formatDisplayDate(from)} to ${formatDisplayDate(to)}`,
    '',
    '## Income',
    `| Category | Amount |`,
    `|----------|--------|`,
    ...data.incomeBreakdown.map((i) => `| ${i.category} | ${formatINR(i.amount)} |`),
    `| **Total Income** | **${formatINR(data.totalIncome)}** |`,
    '',
    '## Cost of Goods Sold',
    `- **Total COGS:** ${formatINR(data.costOfGoodsSold)}`,
    '',
    `## Gross Profit: ${formatINR(data.grossProfit)} (${data.grossMargin.toFixed(1)}% margin)`,
    '',
    '## Operating Expenses',
    `| Category | Amount |`,
    `|----------|--------|`,
    ...data.expenseBreakdown.map((e) => `| ${e.category} | ${formatINR(e.amount)} |`),
    `| **Total Expenses** | **${formatINR(data.operatingExpenses)}** |`,
    '',
    `## Net Profit: ${formatINR(data.netProfit)} (${data.netMargin.toFixed(1)}% margin)`,
    '',
    '## EBITDA',
    `- **EBITDA:** ${formatINR(data.ebitda)}`,
  ];
  return lines.join('\n');
}

function formatMonthlyCashFlow(company: string, month: number, year: number): string {
  const data = generateCashFlowData(month, year);
  const { from, to } = monthToDates(month, year);
  const lines: string[] = [
    `# Cash Flow Statement - ${getMonthName(month)} ${year}`,
    `**Company:** ${company}`,
    `**Period:** ${formatDisplayDate(from)} to ${formatDisplayDate(to)}`,
    '',
    '## Operating Activities',
    ...data.operatingActivities.items.map((i) => `- ${i.description}: ${formatINR(i.amount)}`),
    `**Net Cash from Operations: ${formatINR(data.operatingActivities.total)}**`,
    '',
    '## Investing Activities',
    ...data.investingActivities.items.map((i) => `- ${i.description}: ${formatINR(i.amount)}`),
    `**Net Cash from Investing: ${formatINR(data.investingActivities.total)}**`,
    '',
    '## Financing Activities',
    ...data.financingActivities.items.map((i) => `- ${i.description}: ${formatINR(i.amount)}`),
    `**Net Cash from Financing: ${formatINR(data.financingActivities.total)}**`,
    '',
    '## Cash Position',
    `- **Opening Balance:** ${formatINR(data.openingBalance)}`,
    `- **Net Change:** ${formatINR(data.netCashFlow)}`,
    `- **Closing Balance:** ${formatINR(data.closingBalance)}`,
  ];
  return lines.join('\n');
}

function formatMonthlyExpenseBreakdown(company: string, month: number, year: number): string {
  const pl = generatePLData(month, year);
  const { from, to } = monthToDates(month, year);
  const lines: string[] = [
    `# Expense Breakdown - ${getMonthName(month)} ${year}`,
    `**Company:** ${company}`,
    `**Period:** ${formatDisplayDate(from)} to ${formatDisplayDate(to)}`,
    '',
    '## Expense Categories',
    `| Category | Amount | % of Total |`,
    `|----------|--------|------------|`,
    ...pl.expenseBreakdown.map((e) => {
      const pct = ((e.amount / pl.operatingExpenses) * 100).toFixed(1);
      return `| ${e.category} | ${formatINR(e.amount)} | ${pct}% |`;
    }),
    `| **Total** | **${formatINR(pl.operatingExpenses)}** | **100%** |`,
    '',
    '## Cost of Sales',
    `- **Cost of Goods Sold:** ${formatINR(pl.costOfGoodsSold)}`,
    `- **COGS as % of Sales:** ${((pl.costOfGoodsSold / pl.totalIncome) * 100).toFixed(1)}%`,
    '',
    '## Total Cost Analysis',
    `- **Total Expenses:** ${formatINR(pl.operatingExpenses + pl.costOfGoodsSold)}`,
    `- **Expenses as % of Revenue:** ${(((pl.operatingExpenses + pl.costOfGoodsSold) / pl.totalIncome) * 100).toFixed(1)}%`,
  ];
  return lines.join('\n');
}

function formatMonthlyGST(company: string, month: number, year: number): string {
  const data = generateGSTData(month, year);
  const { from, to } = monthToDates(month, year);
  const lines: string[] = [
    `# GST Summary - ${getMonthName(month)} ${year}`,
    `**Company:** ${company}`,
    `**Period:** ${formatDisplayDate(from)} to ${formatDisplayDate(to)}`,
    '',
    '## Output GST (Tax Collected)',
    `| Tax Type | Amount |`,
    `|----------|--------|`,
    `| IGST | ${formatINR(data.outputGST.igst)} |`,
    `| CGST | ${formatINR(data.outputGST.cgst)} |`,
    `| SGST | ${formatINR(data.outputGST.sgst)} |`,
    `| **Total Output Tax** | **${formatINR(data.outputGST.total)}** |`,
    '',
    '## Input GST (Tax Paid)',
    `| Tax Type | Amount |`,
    `|----------|--------|`,
    `| IGST | ${formatINR(data.inputGST.igst)} |`,
    `| CGST | ${formatINR(data.inputGST.cgst)} |`,
    `| SGST | ${formatINR(data.inputGST.sgst)} |`,
    `| **Total Input Tax** | **${formatINR(data.inputGST.total)}** |`,
    '',
    `## Net GST Liability: ${formatINR(data.netGSTLiability)}`,
    '',
    '## GST Return Status',
    ...data.gstReturns.map((r) => `- **${r.returnType}:** Due by ${r.dueDate} - ${r.status}`),
    '',
    '## HSN Summary',
    `| HSN Code | Description | Taxable Value | Tax Amount |`,
    `|----------|-------------|---------------|------------|`,
    ...data.hsnSummary.map((h) => `| ${h.hsn} | ${h.description} | ${formatINR(h.taxableValue)} | ${formatINR(h.tax)} |`),
  ];
  return lines.join('\n');
}

function formatMonthlyLedgerSummary(company: string, month: number, year: number): string {
  const data = generateLedgerSummaryData(month, year);
  const { from, to } = monthToDates(month, year);
  const lines: string[] = [
    `# Ledger Summary - ${getMonthName(month)} ${year}`,
    `**Company:** ${company}`,
    `**Period:** ${formatDisplayDate(from)} to ${formatDisplayDate(to)}`,
    '',
    '## Ledger-wise Summary',
    `| Ledger | Group | Opening | Debit | Credit | Closing |`,
    `|--------|-------|---------|-------|--------|---------|`,
    ...data.ledgers.map(
      (l) =>
        `| ${l.name} | ${l.group} | ${formatINR(l.opening)} | ${formatINR(l.debit)} | ${formatINR(l.credit)} | ${formatINR(l.closing)} |`
    ),
    '',
    '## Totals',
    `- **Total Debits:** ${formatINR(data.totalDebits)}`,
    `- **Total Credits:** ${formatINR(data.totalCredits)}`,
  ];
  return lines.join('\n');
}

function formatAnnualPL(company: string, year: number): string {
  const data = generateAnnualPLData(year);
  const lines: string[] = [
    `# Annual Profit & Loss Statement - FY ${data.financialYear}`,
    `**Company:** ${company}`,
    '',
    '## Income Statement',
    `| Item | Amount |`,
    `|------|--------|`,
    `| Total Revenue | ${formatINR(data.totalIncome)} |`,
    `| Cost of Goods Sold | ${formatINR(data.costOfGoodsSold)} |`,
    `| **Gross Profit** | **${formatINR(data.grossProfit)}** |`,
    `| Gross Margin | ${data.grossMargin.toFixed(1)}% |`,
    `| Operating Expenses | ${formatINR(data.operatingExpenses)} |`,
    `| **EBITDA** | **${formatINR(data.ebitda)}** |`,
    `| Depreciation | ${formatINR(data.depreciation)} |`,
    `| **EBIT** | **${formatINR(data.ebit)}** |`,
    `| Interest Expense | ${formatINR(data.interestExpense)} |`,
    `| **Profit Before Tax** | **${formatINR(data.pbt)}** |`,
    `| Tax Provision (25%) | ${formatINR(data.taxProvision)} |`,
    `| **Profit After Tax** | **${formatINR(data.pat)}** |`,
    `| Net Margin | ${data.netMargin.toFixed(1)}% |`,
    '',
    '## Quarterly Breakdown',
    `| Quarter | Revenue | Profit |`,
    `|---------|---------|--------|`,
    ...data.quarterlyBreakdown.map((q) => `| ${q.quarter} | ${formatINR(q.income)} | ${formatINR(q.profit)} |`),
  ];
  return lines.join('\n');
}

function formatBalanceSheet(company: string, year: number): string {
  const data = generateBalanceSheetData(year);
  const lines: string[] = [
    `# Balance Sheet`,
    `**Company:** ${company}`,
    `**As on:** ${formatDisplayDate(data.asOnDate)}`,
    '',
    '## Assets',
    '',
    '### Fixed Assets',
    `- Gross Block: ${formatINR(data.assets.fixedAssets.grossBlock)}`,
    `- Less: Accumulated Depreciation: ${formatINR(data.assets.fixedAssets.lessDepreciation)}`,
    `- **Net Block:** ${formatINR(data.assets.fixedAssets.netBlock)}`,
    `- Capital Work in Progress: ${formatINR(data.assets.fixedAssets.capitalWIP)}`,
    `- **Total Fixed Assets:** ${formatINR(data.assets.fixedAssets.total)}`,
    '',
    '### Current Assets',
    `- Inventory: ${formatINR(data.assets.currentAssets.inventory)}`,
    `- Trade Debtors: ${formatINR(data.assets.currentAssets.debtors)}`,
    `- Cash & Bank: ${formatINR(data.assets.currentAssets.cashAndBank)}`,
    `- Loans & Advances: ${formatINR(data.assets.currentAssets.loansAndAdvances)}`,
    `- Other Current Assets: ${formatINR(data.assets.currentAssets.otherCurrentAssets)}`,
    `- **Total Current Assets:** ${formatINR(data.assets.currentAssets.total)}`,
    '',
    `## **Total Assets: ${formatINR(data.assets.totalAssets)}**`,
    '',
    '## Liabilities & Equity',
    '',
    "### Shareholders' Funds",
    `- Share Capital: ${formatINR(data.liabilitiesAndEquity.shareholdersFunds.shareCapital)}`,
    `- Reserves & Surplus: ${formatINR(data.liabilitiesAndEquity.shareholdersFunds.reserves)}`,
    `- **Total Equity:** ${formatINR(data.liabilitiesAndEquity.shareholdersFunds.total)}`,
    '',
    '### Long-Term Liabilities',
    `- Long-Term Borrowings: ${formatINR(data.liabilitiesAndEquity.longTermLiabilities.longTermBorrowings)}`,
    `- Deferred Tax Liability: ${formatINR(data.liabilitiesAndEquity.longTermLiabilities.deferredTaxLiability)}`,
    `- **Total LT Liabilities:** ${formatINR(data.liabilitiesAndEquity.longTermLiabilities.total)}`,
    '',
    '### Current Liabilities',
    `- Short-Term Borrowings: ${formatINR(data.liabilitiesAndEquity.currentLiabilities.shortTermBorrowings)}`,
    `- Trade Creditors: ${formatINR(data.liabilitiesAndEquity.currentLiabilities.creditors)}`,
    `- Other Current Liabilities: ${formatINR(data.liabilitiesAndEquity.currentLiabilities.otherCurrentLiabilities)}`,
    `- **Total Current Liabilities:** ${formatINR(data.liabilitiesAndEquity.currentLiabilities.total)}`,
    '',
    `## **Total Liabilities & Equity: ${formatINR(data.liabilitiesAndEquity.totalLiabilitiesAndEquity)}**`,
    '',
    '## Key Ratios',
    `- **Current Ratio:** ${data.ratios.currentRatio.toFixed(2)}`,
    `- **Debt-to-Equity:** ${data.ratios.debtToEquity.toFixed(2)}`,
    `- **Return on Equity:** ${data.ratios.returnOnEquity.toFixed(1)}%`,
    `- **Return on Assets:** ${data.ratios.returnOnAssets.toFixed(1)}%`,
  ];
  return lines.join('\n');
}

function formatAnnualSalesPurchaseSummary(company: string, year: number): string {
  const currentSales = 7200000;
  const prevSales = 6500000;
  const currentPurchases = 4800000;
  const prevPurchases = 4350000;
  const lines: string[] = [
    `# Annual Sales & Purchase Summary - FY ${year}-${String(year + 1).slice(-2)}`,
    `**Company:** ${company}`,
    '',
    '## Sales Summary',
    `| Period | Sales | Growth |`,
    `|--------|-------|--------|`,
    `| FY ${year}-${String(year + 1).slice(-2)} | ${formatINR(currentSales)} | +${(((currentSales - prevSales) / prevSales) * 100).toFixed(1)}% |`,
    `| FY ${year - 1}-${String(year).slice(-2)} | ${formatINR(prevSales)} | - |`,
    '',
    '## Purchase Summary',
    `| Period | Purchases | Growth |`,
    `|--------|-----------|--------|`,
    `| FY ${year}-${String(year + 1).slice(-2)} | ${formatINR(currentPurchases)} | +${(((currentPurchases - prevPurchases) / prevPurchases) * 100).toFixed(1)}% |`,
    `| FY ${year - 1}-${String(year).slice(-2)} | ${formatINR(prevPurchases)} | - |`,
    '',
    '## Key Metrics',
    `- **Gross Margin:** ${(((currentSales - currentPurchases) / currentSales) * 100).toFixed(1)}%`,
    `- **Purchase-to-Sales Ratio:** ${((currentPurchases / currentSales) * 100).toFixed(1)}%`,
    `- **YoY Sales Growth:** ${(((currentSales - prevSales) / prevSales) * 100).toFixed(1)}%`,
    `- **YoY Purchase Growth:** ${(((currentPurchases - prevPurchases) / prevPurchases) * 100).toFixed(1)}%`,
  ];
  return lines.join('\n');
}

function formatAnnualTaxSummary(company: string, year: number): string {
  const annualSales = 7200000;
  const annualPurchases = 4800000;
  const lines: string[] = [
    `# Annual Tax Summary - FY ${year}-${String(year + 1).slice(-2)}`,
    `**Company:** ${company}`,
    '',
    '## GST Summary',
    `| Tax | Output | Input | Net |`,
    `|-----|--------|-------|-----|`,
    `| IGST | ${formatINR(annualSales * 0.06)} | ${formatINR(annualPurchases * 0.04)} | ${formatINR(annualSales * 0.06 - annualPurchases * 0.04)} |`,
    `| CGST | ${formatINR(annualSales * 0.06)} | ${formatINR(annualPurchases * 0.07)} | ${formatINR(annualSales * 0.06 - annualPurchases * 0.07)} |`,
    `| SGST | ${formatINR(annualSales * 0.06)} | ${formatINR(annualPurchases * 0.07)} | ${formatINR(annualSales * 0.06 - annualPurchases * 0.07)} |`,
    `| **Total** | **${formatINR(annualSales * 0.18)}** | **${formatINR(annualPurchases * 0.18)}** | **${formatINR((annualSales - annualPurchases) * 0.18)}** |`,
    '',
    '## Income Tax',
    `- **Profit Before Tax:** ${formatINR(1200000)}`,
    `- **Tax Rate:** 25%`,
    `- **Tax Provision:** ${formatINR(300000)}`,
    `- **Advance Tax Paid:** ${formatINR(280000)}`,
    `- **Tax Payable:** ${formatINR(20000)}`,
    '',
    '## TDS Summary',
    `- **TDS Deducted (Receivable):** ${formatINR(85000)}`,
    `- **TDS Paid (On Expenses):** ${formatINR(42000)}`,
    `- **Net TDS Position:** ${formatINR(43000)}`,
  ];
  return lines.join('\n');
}

function formatAnnualExpenseIncomeAnalysis(company: string, year: number): string {
  const annualIncome = 7200000;
  const expenseCategories = [
    { name: 'Cost of Goods Sold', amount: 4800000 },
    { name: 'Salaries & Wages', amount: 540000 },
    { name: 'Rent & Utilities', amount: 240000 },
    { name: 'Administrative Expenses', amount: 180000 },
    { name: 'Marketing & Advertising', amount: 144000 },
    { name: 'Depreciation', amount: 120000 },
    { name: 'Finance Costs', amount: 60000 },
    { name: 'Miscellaneous', amount: 96000 },
  ];
  const totalExpenses = expenseCategories.reduce((s, e) => s + e.amount, 0);
  const lines: string[] = [
    `# Annual Expense & Income Analysis - FY ${year}-${String(year + 1).slice(-2)}`,
    `**Company:** ${company}`,
    '',
    '## Income Analysis',
    `- **Sales Revenue:** ${formatINR(annualIncome * 0.9)}`,
    `- **Service Income:** ${formatINR(annualIncome * 0.08)}`,
    `- **Other Income:** ${formatINR(annualIncome * 0.02)}`,
    `- **Total Income:** ${formatINR(annualIncome)}`,
    '',
    '## Expense Analysis',
    `| Category | Amount | % of Revenue |`,
    `|----------|--------|--------------|`,
    ...expenseCategories.map((e) => `| ${e.name} | ${formatINR(e.amount)} | ${((e.amount / annualIncome) * 100).toFixed(1)}% |`),
    `| **Total** | **${formatINR(totalExpenses)}** | **${((totalExpenses / annualIncome) * 100).toFixed(1)}%** |`,
    '',
    '## Profitability',
    `- **Net Profit:** ${formatINR(annualIncome - totalExpenses)}`,
    `- **Net Profit Margin:** ${(((annualIncome - totalExpenses) / annualIncome) * 100).toFixed(1)}%`,
  ];
  return lines.join('\n');
}

function formatFinancialPerformanceComparison(company: string, year: number): string {
  const years = [year - 2, year - 1, year];
  const data = years.map((y) => ({
    year: y,
    revenue: 5800000 + (y - (year - 2)) * 700000,
    grossProfit: 1800000 + (y - (year - 2)) * 200000,
    netProfit: 800000 + (y - (year - 2)) * 200000,
    ebitda: 1100000 + (y - (year - 2)) * 250000,
  }));
  const lines: string[] = [
    `# Financial Performance Comparison`,
    `**Company:** ${company}`,
    `**Comparison Period:** FY ${year - 2}-${String(year - 1).slice(-2)} to FY ${year}-${String(year + 1).slice(-2)}`,
    '',
    '## Revenue Comparison',
    `| FY | Revenue | Growth |`,
    `|----|---------|--------|`,
    ...data.map((d, i) =>
      i === 0
        ? `| ${d.year}-${String(d.year + 1).slice(-2)} | ${formatINR(d.revenue)} | - |`
        : `| ${d.year}-${String(d.year + 1).slice(-2)} | ${formatINR(d.revenue)} | +${(((d.revenue - data[i - 1].revenue) / data[i - 1].revenue) * 100).toFixed(1)}% |`
    ),
    '',
    '## Profitability Comparison',
    `| FY | Gross Profit | Net Profit | EBITDA |`,
    `|----|--------------|------------|--------|`,
    ...data.map((d) => `| ${d.year}-${String(d.year + 1).slice(-2)} | ${formatINR(d.grossProfit)} | ${formatINR(d.netProfit)} | ${formatINR(d.ebitda)} |`),
    '',
    '## Key Trends',
    `- Revenue CAGR: ${(((data[2].revenue / data[0].revenue) ** 0.5 - 1) * 100).toFixed(1)}%`,
    `- Net Profit CAGR: ${(((data[2].netProfit / data[0].netProfit) ** 0.5 - 1) * 100).toFixed(1)}%`,
  ];
  return lines.join('\n');
}

function formatLedgerReport(company: string, ledgerName?: string, groupName?: string, fromDate?: string, toDate?: string): string {
  const data = generateLedgerSummaryData(3, 2025);
  const filtered = ledgerName
    ? data.ledgers.filter((l) => l.name.toLowerCase().includes(ledgerName.toLowerCase()))
    : groupName
      ? data.ledgers.filter((l) => l.group.toLowerCase().includes(groupName.toLowerCase()))
      : data.ledgers;
  const lines: string[] = [
    `# Ledger Report`,
    `**Company:** ${company}`,
    ledgerName ? `**Ledger:** ${ledgerName}` : '',
    groupName ? `**Group:** ${groupName}` : '',
    fromDate ? `**Period:** ${formatDisplayDate(fromDate)} to ${formatDisplayDate(toDate ?? '')}` : '',
    '',
    '## Ledger Details',
    `| Ledger | Group | Opening | Debit | Credit | Closing |`,
    `|--------|-------|---------|-------|--------|---------|`,
    ...filtered.map(
      (l) =>
        `| ${l.name} | ${l.group} | ${formatINR(l.opening)} | ${formatINR(l.debit)} | ${formatINR(l.credit)} | ${formatINR(l.closing)} |`
    ),
  ].filter((l) => l !== '');
  return lines.join('\n');
}

async function formatOutstandingReport(company: string, type: 'receivable' | 'payable' | 'both'): Promise<string> {
  const title = `# Outstanding ${type === 'both' ? 'Receivables & Payables' : type === 'receivable' ? 'Receivables' : 'Payables'} Report`;
  const lines: string[] = [
    title,
    `**Company:** ${company}`,
    `**As on:** ${new Date().toLocaleDateString('en-IN')}`,
    '',
  ];

  let usedRealData = false;

  if (type === 'receivable' || type === 'both') {
    try {
      const xml = await sendTallyRequest(getReceivablesRequest(company));
      const ledgers = await parseLedgerResponse(xml);
      if (ledgers.length > 0) {
        usedRealData = true;
        const total = sumLedgerBalances(ledgers);
        lines.push(
          '## Outstanding Receivables (Sundry Debtors)',
          `| Party | Balance |`,
          `|-------|---------|`,
          ...ledgers.map((l) => `| ${l.name} | ${formatINR(Math.abs(l.closingBalance))} |`),
          `| **Total** | **${formatINR(total)}** |`,
          ''
        );
      }
    } catch {
      // Fall through to demo
    }
  }

  if (type === 'payable' || type === 'both') {
    try {
      const xml = await sendTallyRequest(getPayablesRequest(company));
      const ledgers = await parseLedgerResponse(xml);
      if (ledgers.length > 0) {
        usedRealData = true;
        const total = sumLedgerBalances(ledgers);
        lines.push(
          '## Outstanding Payables (Sundry Creditors)',
          `| Party | Balance |`,
          `|-------|---------|`,
          ...ledgers.map((l) => `| ${l.name} | ${formatINR(Math.abs(l.closingBalance))} |`),
          `| **Total** | **${formatINR(total)}** |`
        );
      }
    } catch {
      // Fall through to demo
    }
  }

  if (usedRealData) {
    return lines.join('\n');
  }

  // Demo fallback
  const data = generateOutstandingData(type);
  const demoLines: string[] = [
    title,
    `**Company:** ${company}`,
    `**As on:** ${new Date().toLocaleDateString('en-IN')}`,
    '',
  ];
  if (type === 'receivable' || type === 'both') {
    const totalRec = data.receivables.reduce((s, r) => s + r.amount, 0);
    demoLines.push(
      '## Outstanding Receivables',
      `| Party | Amount | Days Overdue | Ageing |`,
      `|-------|--------|--------------|--------|`,
      ...data.receivables.map((r) => `| ${r.party} | ${formatINR(r.amount)} | ${r.daysOverdue} | ${r.ageing} |`),
      `| **Total** | **${formatINR(totalRec)}** | | |`,
      ''
    );
  }
  if (type === 'payable' || type === 'both') {
    const totalPay = data.payables.reduce((s, p) => s + p.amount, 0);
    demoLines.push(
      '## Outstanding Payables',
      `| Party | Amount | Days Overdue | Ageing |`,
      `|-------|--------|--------------|--------|`,
      ...data.payables.map((p) => `| ${p.party} | ${formatINR(p.amount)} | ${p.daysOverdue} | ${p.ageing} |`),
      `| **Total** | **${formatINR(totalPay)}** | | |`
    );
  }
  demoLines.push('', `_Note: TallyPrime not reachable — showing demo data._`);
  return demoLines.join('\n');
}

async function formatStockSummary(company: string): Promise<string> {
  try {
    const xml = await sendTallyRequest(getStockSummaryRequest(company));
    const items = await parseStockItemResponse(xml);
    if (items.length > 0) {
      const totalValue = items.reduce((s, i) => s + Math.abs(i.closingValue), 0);
      // Group by parent
      const grouped = new Map<string, typeof items>();
      for (const item of items) {
        const group = item.parent || 'Ungrouped';
        if (!grouped.has(group)) grouped.set(group, []);
        grouped.get(group)!.push(item);
      }
      const lines: string[] = [
        `# Stock Summary`,
        `**Company:** ${company}`,
        `**As on:** ${new Date().toLocaleDateString('en-IN')}`,
        '',
        `## Overview`,
        `- **Total Stock Value:** ${formatINR(totalValue)}`,
        `- **Total Items:** ${items.length}`,
        '',
      ];
      for (const [group, groupItems] of grouped) {
        lines.push(
          `## ${group}`,
          `| Item | Quantity | Value |`,
          `|------|----------|-------|`,
          ...groupItems.map((i) => `| ${i.name} | ${i.closingBalance} | ${formatINR(Math.abs(i.closingValue))} |`),
          ''
        );
      }
      return lines.join('\n');
    }
  } catch {
    // Fall through to demo data
  }

  // Demo fallback
  const data = generateStockData();
  const lines: string[] = [
    `# Stock Summary`,
    `**Company:** ${company}`,
    `**As on:** ${new Date().toLocaleDateString('en-IN')}`,
    '',
    `## Overview`,
    `- **Total Stock Value:** ${formatINR(data.totalStockValue)}`,
    `- **Total Items:** ${data.totalItems}`,
    '',
  ];
  for (const group of data.groups) {
    lines.push(
      `## ${group.group}`,
      `| Item | Unit | Quantity | Rate | Value |`,
      `|------|------|----------|------|-------|`,
      ...group.items.map((i) => `| ${i.name} | ${i.unit} | ${i.quantity} | ${formatINR(i.rate)} | ${formatINR(i.value)} |`),
      ''
    );
  }
  if (data.lowStockItems.length > 0) {
    lines.push(
      '## ⚠️ Low Stock Alert',
      `| Item | Current Stock | Reorder Level |`,
      `|------|---------------|---------------|`,
      ...data.lowStockItems.map((i) => `| ${i.name} | ${i.currentStock} | ${i.reorderLevel} |`)
    );
  }
  lines.push('', `_Note: TallyPrime not reachable — showing demo data._`);
  return lines.join('\n');
}

function formatDaybook(company: string, fromDate: string, toDate: string): string {
  const entries = [
    { date: '01/03/2025', type: 'Sales', number: 'SI/2025/001', party: 'ABC Corporation', amount: 125000 },
    { date: '02/03/2025', type: 'Purchase', number: 'PI/2025/001', party: 'National Supplies', amount: 85000 },
    { date: '03/03/2025', type: 'Payment', number: 'PV/2025/001', party: 'National Supplies', amount: 80000 },
    { date: '04/03/2025', type: 'Receipt', number: 'RV/2025/001', party: 'ABC Corporation', amount: 120000 },
    { date: '05/03/2025', type: 'Journal', number: 'JV/2025/001', party: 'Salary Expense', amount: 45000 },
    { date: '06/03/2025', type: 'Sales', number: 'SI/2025/002', party: 'XYZ Enterprises', amount: 95000 },
  ];
  const lines: string[] = [
    `# Day Book`,
    `**Company:** ${company}`,
    `**Period:** ${formatDisplayDate(fromDate)} to ${formatDisplayDate(toDate)}`,
    '',
    '## Transactions',
    `| Date | Type | Number | Party | Amount |`,
    `|------|------|--------|-------|--------|`,
    ...entries.map((e) => `| ${e.date} | ${e.type} | ${e.number} | ${e.party} | ${formatINR(e.amount)} |`),
    '',
    `**Total Transactions:** ${entries.length}`,
    `**Total Amount:** ${formatINR(entries.reduce((s, e) => s + e.amount, 0))}`,
  ];
  return lines.join('\n');
}

function formatVoucherSummary(company: string, fromDate: string, toDate: string): string {
  const summary = [
    { type: 'Sales', count: 45, totalAmount: 575000 },
    { type: 'Purchase', count: 28, totalAmount: 385000 },
    { type: 'Receipt', count: 38, totalAmount: 520000 },
    { type: 'Payment', count: 32, totalAmount: 410000 },
    { type: 'Journal', count: 12, totalAmount: 85000 },
    { type: 'Contra', count: 8, totalAmount: 120000 },
    { type: 'Debit Note', count: 3, totalAmount: 25000 },
    { type: 'Credit Note', count: 2, totalAmount: 18000 },
  ];
  const lines: string[] = [
    `# Voucher Summary`,
    `**Company:** ${company}`,
    `**Period:** ${formatDisplayDate(fromDate)} to ${formatDisplayDate(toDate)}`,
    '',
    '## Voucher Type Summary',
    `| Voucher Type | Count | Total Amount |`,
    `|--------------|-------|--------------|`,
    ...summary.map((s) => `| ${s.type} | ${s.count} | ${formatINR(s.totalAmount)} |`),
    `| **Total** | **${summary.reduce((t, s) => t + s.count, 0)}** | **${formatINR(summary.reduce((t, s) => t + s.totalAmount, 0))}** |`,
  ];
  return lines.join('\n');
}

function formatPartyWiseBalances(company: string): string {
  const parties = [
    { name: 'ABC Corporation Pvt Ltd', type: 'Customer', balance: 285000, drCr: 'Dr' as const },
    { name: 'XYZ Enterprises', type: 'Customer', balance: 192000, drCr: 'Dr' as const },
    { name: 'Global Traders Ltd', type: 'Customer', balance: 156000, drCr: 'Dr' as const },
    { name: 'National Supplies Co', type: 'Vendor', balance: 185000, drCr: 'Cr' as const },
    { name: 'Metro Distributors', type: 'Vendor', balance: 142000, drCr: 'Cr' as const },
    { name: 'City Wholesalers', type: 'Vendor', balance: 95000, drCr: 'Cr' as const },
  ];
  const customers = parties.filter((p) => p.type === 'Customer');
  const vendors = parties.filter((p) => p.type === 'Vendor');
  const lines: string[] = [
    `# Party-wise Balances`,
    `**Company:** ${company}`,
    `**As on:** ${new Date().toLocaleDateString('en-IN')}`,
    '',
    '## Customer Balances (Debtors)',
    `| Party | Balance | Dr/Cr |`,
    `|-------|---------|-------|`,
    ...customers.map((p) => `| ${p.name} | ${formatINR(p.balance)} | ${p.drCr} |`),
    `| **Total Debtors** | **${formatINR(customers.reduce((s, p) => s + p.balance, 0))}** | Dr |`,
    '',
    '## Vendor Balances (Creditors)',
    `| Party | Balance | Dr/Cr |`,
    `|-------|---------|-------|`,
    ...vendors.map((p) => `| ${p.name} | ${formatINR(p.balance)} | ${p.drCr} |`),
    `| **Total Creditors** | **${formatINR(vendors.reduce((s, p) => s + p.balance, 0))}** | Cr |`,
  ];
  return lines.join('\n');
}

function formatFinancialDashboard(company: string): string {
  const today = new Date().toLocaleDateString('en-IN', { day: '2-digit', month: '2-digit', year: 'numeric' });
  const currentMonth = new Date().getMonth() + 1;
  const currentYear = new Date().getFullYear();
  const sales = generateSalesData(currentMonth, currentYear);
  const pl = generatePLData(currentMonth, currentYear);
  const cash = generateCashFlowData(currentMonth, currentYear);
  const outstanding = generateOutstandingData('both');
  const totalReceivables = outstanding.receivables.reduce((s, r) => s + r.amount, 0);
  const totalPayables = outstanding.payables.reduce((s, p) => s + p.amount, 0);
  const lines: string[] = [
    `# Financial Dashboard`,
    `**Company:** ${company}`,
    `**As on:** ${today}`,
    '',
    '## 📊 Key Performance Indicators',
    '',
    '### Revenue & Profitability',
    `- **Total Sales (MTD):** ${formatINR(sales.totalSales)}`,
    `- **Gross Profit:** ${formatINR(pl.grossProfit)} (${pl.grossMargin.toFixed(1)}% margin)`,
    `- **Net Profit:** ${formatINR(pl.netProfit)} (${pl.netMargin.toFixed(1)}% margin)`,
    `- **EBITDA:** ${formatINR(pl.ebitda)}`,
    '',
    '### Cash Position',
    `- **Cash & Bank Balance:** ${formatINR(cash.closingBalance)}`,
    `- **Operating Cash Flow:** ${formatINR(cash.operatingActivities.total)}`,
    `- **Net Cash Flow (MTD):** ${formatINR(cash.netCashFlow)}`,
    '',
    '### Working Capital',
    `- **Total Receivables:** ${formatINR(totalReceivables)}`,
    `- **Total Payables:** ${formatINR(totalPayables)}`,
    `- **Net Working Capital:** ${formatINR(totalReceivables - totalPayables)}`,
    '',
    '### Growth Metrics',
    `- **MoM Sales Growth:** +${sales.growth.toFixed(1)}%`,
    `- **YoY Projected Growth:** +12.4%`,
    '',
    '## ⚠️ Alerts',
    `- ${outstanding.receivables.filter((r) => r.daysOverdue > 60).length} customer(s) with receivables overdue >60 days`,
    `- ${outstanding.payables.filter((p) => p.daysOverdue > 30).length} vendor(s) with payables overdue >30 days`,
    `- Stock reorder needed for 2 items`,
  ];
  return lines.join('\n');
}

function formatProfitMargins(company: string, month: number, year: number): string {
  const pl = generatePLData(month, year);
  const annualPl = generateAnnualPLData(year);
  const lines: string[] = [
    `# Profit Margins Analysis`,
    `**Company:** ${company}`,
    '',
    `## Monthly Margins - ${getMonthName(month)} ${year}`,
    `| Metric | Amount | Margin % |`,
    `|--------|--------|----------|`,
    `| Revenue | ${formatINR(pl.totalIncome)} | 100.0% |`,
    `| Gross Profit | ${formatINR(pl.grossProfit)} | ${pl.grossMargin.toFixed(1)}% |`,
    `| EBITDA | ${formatINR(pl.ebitda)} | ${((pl.ebitda / pl.totalIncome) * 100).toFixed(1)}% |`,
    `| Net Profit | ${formatINR(pl.netProfit)} | ${pl.netMargin.toFixed(1)}% |`,
    '',
    `## Annual Margins - FY ${year}-${String(year + 1).slice(-2)}`,
    `| Metric | Amount | Margin % |`,
    `|--------|--------|----------|`,
    `| Revenue | ${formatINR(annualPl.totalIncome)} | 100.0% |`,
    `| Gross Profit | ${formatINR(annualPl.grossProfit)} | ${annualPl.grossMargin.toFixed(1)}% |`,
    `| EBITDA | ${formatINR(annualPl.ebitda)} | ${((annualPl.ebitda / annualPl.totalIncome) * 100).toFixed(1)}% |`,
    `| PAT | ${formatINR(annualPl.pat)} | ${annualPl.netMargin.toFixed(1)}% |`,
    '',
    '## Interpretation',
    `- Gross margin of ${pl.grossMargin.toFixed(1)}% is ${pl.grossMargin > 30 ? 'healthy' : 'below industry average'}`,
    `- Net margin of ${pl.netMargin.toFixed(1)}% indicates ${pl.netMargin > 10 ? 'good' : 'moderate'} profitability`,
  ];
  return lines.join('\n');
}

function formatGrowthComparison(company: string, year: number): string {
  const months = Array.from({ length: 12 }, (_, i) => i + 1);
  const currentYearData = months.map((m) => ({ month: m, sales: generateSalesData(m, year).totalSales }));
  const prevYearData = months.map((m) => ({ month: m, sales: generateSalesData(m, year - 1).totalSales * 0.88 }));
  const lines: string[] = [
    `# Growth Comparison Analysis`,
    `**Company:** ${company}`,
    '',
    '## Month-on-Month Growth (Current Year)',
    `| Month | Sales | MoM Growth |`,
    `|-------|-------|------------|`,
    ...currentYearData.map((d, i) => {
      const prev = i > 0 ? currentYearData[i - 1].sales : d.sales;
      const growth = i > 0 ? (((d.sales - prev) / prev) * 100).toFixed(1) : '-';
      return `| ${getMonthName(d.month)} | ${formatINR(d.sales)} | ${growth !== '-' ? `+${growth}%` : '-'} |`;
    }),
    '',
    '## Year-on-Year Comparison',
    `| Month | FY ${year} | FY ${year - 1} | YoY Growth |`,
    `|-------|--------|--------|------------|`,
    ...currentYearData.map((d, i) => {
      const prevYr = prevYearData[i].sales;
      const growth = (((d.sales - prevYr) / prevYr) * 100).toFixed(1);
      return `| ${getMonthName(d.month)} | ${formatINR(d.sales)} | ${formatINR(prevYr)} | +${growth}% |`;
    }),
    '',
    '## Summary',
    `- **Annual YoY Growth:** +${(((currentYearData.reduce((s, d) => s + d.sales, 0) - prevYearData.reduce((s, d) => s + d.sales, 0)) / prevYearData.reduce((s, d) => s + d.sales, 0)) * 100).toFixed(1)}%`,
  ];
  return lines.join('\n');
}

function formatCustomAnalysis(
  company: string,
  analysisType: string,
  period: string,
  year: number,
  month?: number
): string {
  const periodLabel = period === 'monthly' && month ? `${getMonthName(month)} ${year}` : period === 'annual' ? `FY ${year}-${String(year + 1).slice(-2)}` : `Q${Math.ceil((month ?? 1) / 3)} ${year}`;
  const lines: string[] = [
    `# Custom Financial Analysis - ${analysisType.charAt(0).toUpperCase() + analysisType.slice(1)}`,
    `**Company:** ${company}`,
    `**Period:** ${periodLabel}`,
    '',
  ];
  switch (analysisType) {
    case 'revenue':
      lines.push(
        '## Revenue Analysis',
        `- **Product Sales:** ${formatINR(520000)}`,
        `- **Service Revenue:** ${formatINR(80000)}`,
        `- **Other Income:** ${formatINR(15000)}`,
        `- **Total Revenue:** ${formatINR(615000)}`,
        '',
        '## Revenue Mix',
        `- Product: 84.6%  |  Services: 13.0%  |  Other: 2.4%`
      );
      break;
    case 'expense':
      lines.push(
        '## Expense Analysis',
        `- **Direct Costs:** ${formatINR(380000)} (61.8% of revenue)`,
        `- **Indirect Costs:** ${formatINR(235000)} (38.2% of revenue)`,
        `- **Cost per Unit Sold:** ${formatINR(850)}`,
        '',
        '## Cost Drivers',
        `- Raw Materials: 45%  |  Labour: 25%  |  Overhead: 30%`
      );
      break;
    case 'profitability':
      lines.push(
        '## Profitability Analysis',
        `- **Gross Margin:** 38.2%`,
        `- **EBITDA Margin:** 22.1%`,
        `- **Net Margin:** 16.5%`,
        `- **ROE:** 18.5%  |  **ROA:** 9.4%  |  **ROCE:** 15.2%`
      );
      break;
    case 'liquidity':
      lines.push(
        '## Liquidity Analysis',
        `- **Current Ratio:** 1.93 (Healthy >1.5)`,
        `- **Quick Ratio:** 1.45 (Healthy >1.0)`,
        `- **Cash Ratio:** 0.33`,
        `- **Operating Cash Flow Ratio:** 0.89`
      );
      break;
    case 'efficiency':
      lines.push(
        '## Efficiency Analysis',
        `- **Inventory Turnover:** 7.4x`,
        `- **Debtor Days:** 48 days`,
        `- **Creditor Days:** 35 days`,
        `- **Asset Turnover:** 1.31x`,
        `- **Cash Conversion Cycle:** 60 days`
      );
      break;
    default:
      lines.push('Analysis type not recognized.');
  }
  return lines.join('\n');
}

async function formatCompanyList(): Promise<string> {
  try {
    const xml = await sendTallyRequest(getCompanyListRequest());
    const companies = await parseCompanyListResponse(xml);
    if (companies.length > 0) {
      const lines: string[] = [
        '# Available Tally Companies',
        '',
        `**Current Default:** ${connectionState.defaultCompany}`,
        '',
        '## Companies',
        `| # | Company Name | GUID | Books From | Last Voucher |`,
        `|---|--------------|------|------------|--------------|`,
        ...companies.map(
          (c, i) =>
            `| ${i + 1} | ${c.name} | ${c.guid || 'N/A'} | ${c.booksFrom || 'N/A'} | ${c.lastVoucherDate || 'N/A'} |`
        ),
      ];
      return lines.join('\n');
    }
  } catch {
    // Fall through to demo data if Tally is unavailable
  }

  // Demo fallback
  const companies: TallyCompany[] = [
    {
      name: 'Sample Company Ltd.',
      guid: 'GUID-001-SAMPLE',
      startingFrom: '01-04-2020',
      booksFrom: '01-04-2020',
      lastVoucherDate: '06-03-2025',
    },
    {
      name: 'Demo Enterprises Pvt Ltd',
      guid: 'GUID-002-DEMO',
      startingFrom: '01-04-2019',
      booksFrom: '01-04-2019',
      lastVoucherDate: '28-02-2025',
    },
    {
      name: 'Test Trading Co',
      guid: 'GUID-003-TEST',
      startingFrom: '01-04-2022',
      booksFrom: '01-04-2022',
      lastVoucherDate: '05-03-2025',
    },
  ];
  const lines: string[] = [
    '# Available Tally Companies',
    '',
    `**Current Default:** ${connectionState.defaultCompany}`,
    '',
    '## Companies',
    `| # | Company Name | GUID | Books From | Last Voucher |`,
    `|---|--------------|------|------------|--------------|`,
    ...companies.map(
      (c, i) =>
        `| ${i + 1} | ${c.name} | ${c.guid} | ${formatDisplayDate(c.booksFrom)} | ${formatDisplayDate(c.lastVoucherDate)} |`
    ),
    '',
    `_Note: TallyPrime not reachable — showing demo data. Ensure TallyPrime is running and a company is loaded._`,
  ];
  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Tool Definitions
// ---------------------------------------------------------------------------

const TOOLS: Tool[] = [
  // ── Monthly Reports ──────────────────────────────────────────────────────
  {
    name: 'get_monthly_sales_summary',
    description: 'Get monthly sales summary with total sales, top customers, and month-on-month trends',
    inputSchema: {
      type: 'object',
      properties: {
        month: { type: 'number', description: 'Month number (1-12)', minimum: 1, maximum: 12 },
        year: { type: 'number', description: 'Four-digit year', minimum: 2000, maximum: 2100 },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['month', 'year'],
    },
  },
  {
    name: 'get_monthly_purchase_summary',
    description: 'Get monthly purchase summary with total purchases, top vendors, and trends',
    inputSchema: {
      type: 'object',
      properties: {
        month: { type: 'number', description: 'Month number (1-12)', minimum: 1, maximum: 12 },
        year: { type: 'number', description: 'Four-digit year', minimum: 2000, maximum: 2100 },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['month', 'year'],
    },
  },
  {
    name: 'get_monthly_profit_loss',
    description: 'Get monthly profit and loss statement with gross profit, operating expenses, and net profit',
    inputSchema: {
      type: 'object',
      properties: {
        month: { type: 'number', description: 'Month number (1-12)', minimum: 1, maximum: 12 },
        year: { type: 'number', description: 'Four-digit year', minimum: 2000, maximum: 2100 },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['month', 'year'],
    },
  },
  {
    name: 'get_monthly_cash_flow',
    description: 'Get monthly cash flow statement showing operating, investing, and financing activities',
    inputSchema: {
      type: 'object',
      properties: {
        month: { type: 'number', description: 'Month number (1-12)', minimum: 1, maximum: 12 },
        year: { type: 'number', description: 'Four-digit year', minimum: 2000, maximum: 2100 },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['month', 'year'],
    },
  },
  {
    name: 'get_monthly_expense_breakdown',
    description: 'Get detailed monthly expense categorization with percentage breakdown',
    inputSchema: {
      type: 'object',
      properties: {
        month: { type: 'number', description: 'Month number (1-12)', minimum: 1, maximum: 12 },
        year: { type: 'number', description: 'Four-digit year', minimum: 2000, maximum: 2100 },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['month', 'year'],
    },
  },
  {
    name: 'get_monthly_gst_summary',
    description: 'Get monthly GST summary with input/output tax, net liability, and return status',
    inputSchema: {
      type: 'object',
      properties: {
        month: { type: 'number', description: 'Month number (1-12)', minimum: 1, maximum: 12 },
        year: { type: 'number', description: 'Four-digit year', minimum: 2000, maximum: 2100 },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['month', 'year'],
    },
  },
  {
    name: 'get_monthly_ledger_summary',
    description: 'Get monthly ledger-wise transaction summary with opening/closing balances',
    inputSchema: {
      type: 'object',
      properties: {
        month: { type: 'number', description: 'Month number (1-12)', minimum: 1, maximum: 12 },
        year: { type: 'number', description: 'Four-digit year', minimum: 2000, maximum: 2100 },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['month', 'year'],
    },
  },
  // ── Annual Reports ────────────────────────────────────────────────────────
  {
    name: 'get_annual_profit_loss',
    description: 'Get annual profit and loss statement including EBITDA, depreciation, and PAT',
    inputSchema: {
      type: 'object',
      properties: {
        year: { type: 'number', description: 'Financial year start (e.g., 2024 for FY 2024-25)', minimum: 2000, maximum: 2100 },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['year'],
    },
  },
  {
    name: 'get_balance_sheet',
    description: 'Get balance sheet showing assets, liabilities, and equity with key ratios',
    inputSchema: {
      type: 'object',
      properties: {
        year: { type: 'number', description: 'Financial year (e.g., 2024 for 31-Mar-2025 balance sheet)', minimum: 2000, maximum: 2100 },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['year'],
    },
  },
  {
    name: 'get_annual_sales_purchase_summary',
    description: 'Get annual sales and purchase summary with year-on-year comparisons',
    inputSchema: {
      type: 'object',
      properties: {
        year: { type: 'number', description: 'Financial year start', minimum: 2000, maximum: 2100 },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['year'],
    },
  },
  {
    name: 'get_annual_tax_summary',
    description: 'Get annual tax summary covering GST, income tax, and TDS',
    inputSchema: {
      type: 'object',
      properties: {
        year: { type: 'number', description: 'Financial year start', minimum: 2000, maximum: 2100 },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['year'],
    },
  },
  {
    name: 'get_annual_expense_income_analysis',
    description: 'Get detailed annual expense and income analysis with category breakdown',
    inputSchema: {
      type: 'object',
      properties: {
        year: { type: 'number', description: 'Financial year start', minimum: 2000, maximum: 2100 },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['year'],
    },
  },
  {
    name: 'get_financial_performance_comparison',
    description: 'Get multi-year financial performance comparison with CAGR analysis',
    inputSchema: {
      type: 'object',
      properties: {
        year: { type: 'number', description: 'Most recent financial year for comparison', minimum: 2002, maximum: 2100 },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['year'],
    },
  },
  // ── Operational Reports ───────────────────────────────────────────────────
  {
    name: 'get_ledger_report',
    description: 'Get detailed ledger report for a specific ledger or group with transaction details',
    inputSchema: {
      type: 'object',
      properties: {
        ledger_name: { type: 'string', description: 'Specific ledger name to filter by' },
        group_name: { type: 'string', description: 'Ledger group name to filter by' },
        from_date: { type: 'string', description: 'Start date (DD-MM-YYYY or YYYY-MM-DD)' },
        to_date: { type: 'string', description: 'End date (DD-MM-YYYY or YYYY-MM-DD)' },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
    },
  },
  {
    name: 'get_outstanding_receivables_payables',
    description: 'Get outstanding receivables and/or payables with ageing analysis',
    inputSchema: {
      type: 'object',
      properties: {
        type: {
          type: 'string',
          enum: ['receivable', 'payable', 'both'],
          description: 'Type of outstanding: receivable, payable, or both',
        },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['type'],
    },
  },
  {
    name: 'get_stock_summary',
    description: 'Get stock summary with quantities, values, and low stock alerts',
    inputSchema: {
      type: 'object',
      properties: {
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
    },
  },
  {
    name: 'get_daybook',
    description: 'Get day book showing all transactions for a specified period',
    inputSchema: {
      type: 'object',
      properties: {
        from_date: { type: 'string', description: 'Start date (DD-MM-YYYY or YYYY-MM-DD)' },
        to_date: { type: 'string', description: 'End date (DD-MM-YYYY or YYYY-MM-DD)' },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['from_date', 'to_date'],
    },
  },
  {
    name: 'get_voucher_summary',
    description: 'Get voucher summary by type for a specified period',
    inputSchema: {
      type: 'object',
      properties: {
        from_date: { type: 'string', description: 'Start date (DD-MM-YYYY or YYYY-MM-DD)' },
        to_date: { type: 'string', description: 'End date (DD-MM-YYYY or YYYY-MM-DD)' },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['from_date', 'to_date'],
    },
  },
  {
    name: 'get_party_wise_balances',
    description: 'Get customer and supplier balances with outstanding amounts',
    inputSchema: {
      type: 'object',
      properties: {
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
    },
  },
  // ── Dashboard & Analytics ─────────────────────────────────────────────────
  {
    name: 'get_financial_dashboard',
    description: 'Get comprehensive financial dashboard with KPIs, alerts, and key metrics',
    inputSchema: {
      type: 'object',
      properties: {
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
    },
  },
  {
    name: 'calculate_profit_margins',
    description: 'Calculate gross, net, and operating profit margins for a given period',
    inputSchema: {
      type: 'object',
      properties: {
        month: { type: 'number', description: 'Month number (1-12)', minimum: 1, maximum: 12 },
        year: { type: 'number', description: 'Four-digit year', minimum: 2000, maximum: 2100 },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['month', 'year'],
    },
  },
  {
    name: 'calculate_growth_comparison',
    description: 'Calculate month-on-month and year-on-year growth percentages',
    inputSchema: {
      type: 'object',
      properties: {
        year: { type: 'number', description: 'Financial year for comparison', minimum: 2001, maximum: 2100 },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['year'],
    },
  },
  {
    name: 'calculate_custom_analysis',
    description: 'Perform custom financial analysis: revenue, expense, profitability, liquidity, or efficiency',
    inputSchema: {
      type: 'object',
      properties: {
        analysis_type: {
          type: 'string',
          enum: ['revenue', 'expense', 'profitability', 'liquidity', 'efficiency'],
          description: 'Type of financial analysis',
        },
        period: {
          type: 'string',
          enum: ['monthly', 'quarterly', 'annual'],
          description: 'Analysis period',
        },
        year: { type: 'number', description: 'Financial year', minimum: 2000, maximum: 2100 },
        month: { type: 'number', description: 'Month (required for monthly period)', minimum: 1, maximum: 12 },
        company: { type: 'string', description: 'Company name (uses default if not specified)' },
      },
      required: ['analysis_type', 'period', 'year'],
    },
  },
  // ── Utility Tools ─────────────────────────────────────────────────────────
  {
    name: 'get_company_list',
    description: 'Get list of available companies in TallyPrime',
    inputSchema: {
      type: 'object',
      properties: {},
    },
  },
  {
    name: 'set_default_company',
    description: 'Set the default company to use for subsequent operations',
    inputSchema: {
      type: 'object',
      properties: {
        company: { type: 'string', description: 'Company name to set as default' },
      },
      required: ['company'],
    },
  },
];

// ---------------------------------------------------------------------------
// MCP Server Factory
// ---------------------------------------------------------------------------

/**
 * Creates and configures a new MCP Server instance with all tool handlers.
 * Called once for stdio mode, or per-session in HTTP mode.
 */
function createMcpServer(): Server {
  const srv = new Server(
    {
      name: 'tally-mcp-server',
      version: '2.0.0',
    },
    {
      capabilities: {
        tools: {},
      },
    }
  );

  // Handle list tools request
  srv.setRequestHandler(ListToolsRequestSchema, async () => {
    return { tools: TOOLS };
  });

  // Handle call tool request
  srv.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    switch (name) {
      // ── Monthly Reports ──────────────────────────────────────────────────
      case 'get_monthly_sales_summary': {
        const parsed = MonthYearSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = await formatMonthlySalesSummary(company, parsed.month, parsed.year);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'get_monthly_purchase_summary': {
        const parsed = MonthYearSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = await formatMonthlyPurchaseSummary(company, parsed.month, parsed.year);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'get_monthly_profit_loss': {
        const parsed = MonthYearSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatMonthlyPL(company, parsed.month, parsed.year);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'get_monthly_cash_flow': {
        const parsed = MonthYearSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatMonthlyCashFlow(company, parsed.month, parsed.year);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'get_monthly_expense_breakdown': {
        const parsed = MonthYearSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatMonthlyExpenseBreakdown(company, parsed.month, parsed.year);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'get_monthly_gst_summary': {
        const parsed = MonthYearSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatMonthlyGST(company, parsed.month, parsed.year);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'get_monthly_ledger_summary': {
        const parsed = MonthYearSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatMonthlyLedgerSummary(company, parsed.month, parsed.year);
        return { content: [{ type: 'text', text: report }] };
      }

      // ── Annual Reports ───────────────────────────────────────────────────
      case 'get_annual_profit_loss': {
        const parsed = YearSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatAnnualPL(company, parsed.year);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'get_balance_sheet': {
        const parsed = YearSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatBalanceSheet(company, parsed.year);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'get_annual_sales_purchase_summary': {
        const parsed = YearSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatAnnualSalesPurchaseSummary(company, parsed.year);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'get_annual_tax_summary': {
        const parsed = YearSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatAnnualTaxSummary(company, parsed.year);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'get_annual_expense_income_analysis': {
        const parsed = YearSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatAnnualExpenseIncomeAnalysis(company, parsed.year);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'get_financial_performance_comparison': {
        const parsed = YearSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatFinancialPerformanceComparison(company, parsed.year);
        return { content: [{ type: 'text', text: report }] };
      }

      // ── Operational Reports ──────────────────────────────────────────────
      case 'get_ledger_report': {
        const parsed = LedgerQuerySchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatLedgerReport(
          company,
          parsed.ledger_name,
          parsed.group_name,
          parsed.from_date,
          parsed.to_date
        );
        return { content: [{ type: 'text', text: report }] };
      }

      case 'get_outstanding_receivables_payables': {
        const parsed = z
          .object({
            type: z.enum(['receivable', 'payable', 'both']),
            company: z.string().optional(),
          })
          .parse(args);
        const company = resolveCompany(parsed.company);
        const report = await formatOutstandingReport(company, parsed.type);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'get_stock_summary': {
        const parsed = z.object({ company: z.string().optional() }).parse(args ?? {});
        const company = resolveCompany(parsed.company);
        const report = await formatStockSummary(company);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'get_daybook': {
        const parsed = DateRangeSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatDaybook(company, parsed.from_date, parsed.to_date);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'get_voucher_summary': {
        const parsed = DateRangeSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatVoucherSummary(company, parsed.from_date, parsed.to_date);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'get_party_wise_balances': {
        const parsed = z.object({ company: z.string().optional() }).parse(args ?? {});
        const company = resolveCompany(parsed.company);
        const report = formatPartyWiseBalances(company);
        return { content: [{ type: 'text', text: report }] };
      }

      // ── Dashboard & Analytics ────────────────────────────────────────────
      case 'get_financial_dashboard': {
        const parsed = z.object({ company: z.string().optional() }).parse(args ?? {});
        const company = resolveCompany(parsed.company);
        const report = formatFinancialDashboard(company);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'calculate_profit_margins': {
        const parsed = MonthYearSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatProfitMargins(company, parsed.month, parsed.year);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'calculate_growth_comparison': {
        const parsed = YearSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatGrowthComparison(company, parsed.year);
        return { content: [{ type: 'text', text: report }] };
      }

      case 'calculate_custom_analysis': {
        const parsed = AnalysisSchema.parse(args);
        const company = resolveCompany(parsed.company);
        const report = formatCustomAnalysis(company, parsed.analysis_type, parsed.period, parsed.year, parsed.month);
        return { content: [{ type: 'text', text: report }] };
      }

      // ── Utility Tools ────────────────────────────────────────────────────
      case 'get_company_list': {
        const report = await formatCompanyList();
        return { content: [{ type: 'text', text: report }] };
      }

      case 'set_default_company': {
        const parsed = CompanySchema.parse(args);
        connectionState.defaultCompany = parsed.company;
        return {
          content: [
            {
              type: 'text',
              text: `✅ Default company set to: **${parsed.company}**\n\nAll subsequent operations will use this company unless overridden.`,
            },
          ],
        };
      }

      default:
        return {
          content: [
            {
              type: 'text',
              text: `❌ Unknown tool: ${name}\n\nUse \`get_company_list\` to see available tools or refer to the documentation.`,
            },
          ],
          isError: true,
        };
    }
  } catch (error) {
    if (error instanceof z.ZodError) {
      const issues = error.errors.map((e) => `- ${e.path.join('.')}: ${e.message}`).join('\n');
      return {
        content: [
          {
            type: 'text',
            text: `❌ Invalid input parameters:\n\n${issues}`,
          },
        ],
        isError: true,
      };
    }
    const message = error instanceof Error ? error.message : 'An unexpected error occurred';
    return {
      content: [
        {
          type: 'text',
          text: `❌ Error executing ${name}: ${message}`,
        },
      ],
      isError: true,
    };
  }
  });

  return srv;
}

// ---------------------------------------------------------------------------
// Start Server
// ---------------------------------------------------------------------------

/** Names of all registered tools (used for the GET / info endpoint). */
const TOOL_NAMES = TOOLS.map((t) => t.name);

async function main() {
  const transportMode = process.env.TRANSPORT ?? 'http';

  if (transportMode === 'stdio') {
    // Legacy stdio mode — keeps backward compatibility with CLI/local dev
    const transport = new StdioServerTransport();
    const srv = createMcpServer();
    await srv.connect(transport);
    console.error('Tally MCP Server running on stdio');
    return;
  }

  // HTTP mode (default) — required for Copilot Studio / StreamableHTTP
  const app = express();
  app.use(express.json());

  // CORS headers so Copilot Studio (browser-based) can connect
  app.use((_req: Request, res: Response, next: NextFunction) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, Mcp-Session-Id');
    res.setHeader('Access-Control-Expose-Headers', 'Mcp-Session-Id');
    if (_req.method === 'OPTIONS') {
      res.sendStatus(200);
      return;
    }
    next();
  });

  // Optional API-key auth for /mcp
  const apiKey = process.env.MCP_API_KEY;
  const mcpAuthMiddleware = (req: Request, res: Response, next: NextFunction) => {
    if (!apiKey) {
      // No key configured → allow unauthenticated access (for Copilot Studio)
      return next();
    }
    const authHeader = req.headers['authorization'];
    if (!authHeader || authHeader !== `Bearer ${apiKey}`) {
      res.status(401).json({ error: 'Unauthorized' });
      return;
    }
    next();
  };

  // Per-session transport store
  const sessions = new Map<string, StreamableHTTPServerTransport>();

  // POST /mcp — main MCP endpoint (Copilot Studio uses this)
  app.post('/mcp', mcpAuthMiddleware, async (req: Request, res: Response) => {
    const sessionId = req.headers['mcp-session-id'] as string | undefined;
    let transport: StreamableHTTPServerTransport;

    if (sessionId && sessions.has(sessionId)) {
      transport = sessions.get(sessionId)!;
    } else {
      // Create a new transport + server instance for this session
      transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => randomUUID(),
        onsessioninitialized: (sid) => {
          sessions.set(sid, transport);
        },
      });
      transport.onclose = () => {
        if (transport.sessionId) sessions.delete(transport.sessionId);
      };
      const sessionServer = createMcpServer();
      await sessionServer.connect(transport);
    }
    await transport.handleRequest(req, res, req.body);
  });

  // GET /mcp — SSE stream for existing sessions
  app.get('/mcp', mcpAuthMiddleware, async (req: Request, res: Response) => {
    const sessionId = req.headers['mcp-session-id'] as string | undefined;
    const transport = sessionId ? sessions.get(sessionId) : undefined;
    if (!transport) {
      res.status(400).json({ error: 'Invalid or missing session ID' });
      return;
    }
    await transport.handleRequest(req, res);
  });

  // DELETE /mcp — close a session
  app.delete('/mcp', mcpAuthMiddleware, async (req: Request, res: Response) => {
    const sessionId = req.headers['mcp-session-id'] as string | undefined;
    const transport = sessionId ? sessions.get(sessionId) : undefined;
    if (!transport) {
      res.status(400).json({ error: 'Invalid or missing session ID' });
      return;
    }
    await transport.handleRequest(req, res);
    sessions.delete(sessionId!);
  });

  // GET / — server info
  app.get('/', (_req: Request, res: Response) => {
    res.json({
      name: 'Tally MCP Server',
      version: '2.0.0',
      tools: TOOL_NAMES,
    });
  });

  // GET /health — health check
  app.get('/health', (_req: Request, res: Response) => {
    res.json({ status: 'ok', server: 'Tally MCP Server', version: '2.0.0' });
  });

  const port = parseInt(process.env.PORT ?? '3000', 10);
  app.listen(port, '0.0.0.0', () => {
    console.error(`Tally MCP Server running on http://0.0.0.0:${port}`);
    console.error(`MCP endpoint: http://0.0.0.0:${port}/mcp`);
    console.error(`Auth: ${apiKey ? 'API key required' : 'No auth (open)'}`);
  });
}

main().catch((error: unknown) => {
  console.error('Fatal error starting server:', error);
  process.exit(1);
});
