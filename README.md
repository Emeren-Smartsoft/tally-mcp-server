# Tally MCP Server

A comprehensive **Model Context Protocol (MCP) server** for integrating **TallyPrime** accounting software with Copilot Studio agents. Enables automated financial reporting, analysis, and calculations through a standardized interface.

---

## Features

- **27+ Financial Tools** covering monthly, annual, operational, and multi-company reports
- **Copilot Studio Ready** — designed for seamless agent integration
- **Indian Accounting Support** — ₹ currency, GST, Indian financial year (Apr–Mar)
- **Multi-Company Support** — query and aggregate data from up to 5 companies in a single request
- **Multiple Connection Methods** — ODBC (primary), XML API, and file access
- **TypeScript** with full type safety and Zod input validation
- **Demo Mode** — works out of the box without a live Tally installation
- **Docker Support** — production-ready containerization
- **Request Queue & Retry** — prevents Tally from hanging under concurrent load; automatically retries on transient failures

---

## Available Tools (27 total)

### Monthly Reports (7 tools)
| Tool | Description |
|------|-------------|
| `get_monthly_sales_summary` | Sales summary with top customers and trends |
| `get_monthly_purchase_summary` | Purchase summary with top vendors |
| `get_monthly_profit_loss` | Monthly P&L statement |
| `get_monthly_cash_flow` | Cash flow with operating/investing/financing activities |
| `get_monthly_expense_breakdown` | Detailed expense categorization |
| `get_monthly_gst_summary` | GST input/output tax and net liability |
| `get_monthly_ledger_summary` | Ledger-wise transaction summary |

### Annual Reports (6 tools)
| Tool | Description |
|------|-------------|
| `get_annual_profit_loss` | Full annual P&L with EBITDA and PAT |
| `get_balance_sheet` | Balance sheet with key ratios |
| `get_annual_sales_purchase_summary` | Annual sales/purchase with YoY comparison |
| `get_annual_tax_summary` | GST, income tax, and TDS summary |
| `get_annual_expense_income_analysis` | Detailed expense/income analysis |
| `get_financial_performance_comparison` | Multi-year performance with CAGR |

### Operational Reports (6 tools)
| Tool | Description |
|------|-------------|
| `get_ledger_report` | Ledger/group report with transaction details |
| `get_outstanding_receivables_payables` | Outstanding balances with ageing |
| `get_stock_summary` | Stock quantities, values, and low-stock alerts |
| `get_daybook` | Day book for a specified period |
| `get_voucher_summary` | Voucher summary by type |
| `get_party_wise_balances` | Customer/supplier balances |

### Dashboard & Analytics (4 tools)
| Tool | Description |
|------|-------------|
| `get_financial_dashboard` | Comprehensive KPI dashboard |
| `calculate_profit_margins` | Gross, net, and operating margins |
| `calculate_growth_comparison` | MoM and YoY growth percentages |
| `calculate_custom_analysis` | Custom analysis: revenue/expense/profitability/liquidity/efficiency |

### Utility Tools (2 tools)
| Tool | Description |
|------|-------------|
| `get_company_list` | List available Tally companies |
| `set_default_company` | Set default company for operations |

### Multi-Company Tools (2 tools)
| Tool | Description |
|------|-------------|
| `get_multi_company_sales_summary` | Monthly sales aggregated across multiple companies |
| `get_multi_company_report` | Generic multi-company report (sales, purchases, outstanding, stock) |

> **Multi-company tools** accept a `companies` array (1–5 names) and return a company-wise breakdown plus an aggregate total. One failing company does not block others.

---

## Prerequisites

- **Node.js** >= 18.0.0
- **TallyPrime** (for live data; demo mode works without it)
- **ODBC Driver for Tally** (optional, for ODBC connection)

---

## Installation

```bash
# Clone the repository
git clone https://github.com/Emeren-Smartsoft/tally-mcp-server.git
cd tally-mcp-server

# Install dependencies
npm install

# Build TypeScript
npm run build
```

---

## Configuration

Edit `tally-config.yml` to configure your Tally connection:

```yaml
connection:
  method: odbc          # odbc | xml_api | file
  odbc:
    dsn: "TallyODBC"
  xml_api:
    host: "localhost"
    port: 9000

defaults:
  company: "Your Company Name"
```

---

## Usage

### Start the server
```bash
# Production (after build)
npm start

# Development (with hot reload)
npm run dev
```

### Copilot Studio Integration

Add to your Copilot Studio agent configuration:

```json
{
  "mcpServers": {
    "tally": {
      "command": "node",
      "args": ["path/to/dist/tally-mcp-server.js"]
    }
  }
}
```

---

## Usage Examples

### Get Monthly Sales Summary
```json
{
  "tool": "get_monthly_sales_summary",
  "arguments": {
    "month": 3,
    "year": 2025,
    "company": "Sample Company Ltd."
  }
}
```

**Sample Output:**
```
# Sales Summary - March 2025
**Company:** Sample Company Ltd.
**Period:** 01/03/2025 to 31/03/2025

## Key Metrics
- **Total Sales (incl. Tax):** ₹5,37,500
- **Net Sales (excl. Tax):** ₹4,40,950
- **Number of Transactions:** 43
- **Month-on-Month Growth:** +8.7%

## Top Customers
1. **ABC Corporation Pvt Ltd:** ₹1,18,250
...
```

### Get Financial Dashboard
```json
{
  "tool": "get_financial_dashboard",
  "arguments": {}
}
```

### Get Balance Sheet
```json
{
  "tool": "get_balance_sheet",
  "arguments": {
    "year": 2024
  }
}
```

### Calculate Custom Analysis
```json
{
  "tool": "calculate_custom_analysis",
  "arguments": {
    "analysis_type": "profitability",
    "period": "annual",
    "year": 2024
  }
}
```

---

## Docker Deployment

### Build and run with Docker
```bash
docker build -t tally-mcp-server .
docker run -it tally-mcp-server
```

### Run with Docker Compose
```bash
# Production
docker-compose up tally-mcp-server

# With test database
docker-compose --profile testing up
```

---

## Development

```bash
# Type check only (no emit)
npm run typecheck

# Lint
npm run lint

# Fix lint issues
npm run lint:fix

# Format with Prettier
npm run format
```

---

## Project Structure

```
/
├── src/
│   └── tally-mcp-server.ts    # Main MCP server implementation
├── dist/                      # Compiled JavaScript (after build)
├── package.json               # Dependencies and scripts
├── tsconfig.json              # TypeScript configuration
├── .eslintrc.json             # ESLint configuration
├── .prettierrc                # Prettier configuration
├── tally-config.yml           # Tally connection configuration
├── Dockerfile                 # Docker containerization
├── docker-compose.yml         # Docker Compose deployment
└── .gitignore
```

---

## Architecture

The server uses a **layered architecture**:

1. **MCP Layer** — `@modelcontextprotocol/sdk` handles protocol communication
2. **Tool Layer** — 25 named tools with Zod-validated inputs
3. **Report Layer** — Markdown-formatted report generators
4. **Data Layer** — ODBC/XML API queries (demo data in standalone mode)

### Connection Methods (in priority order)
1. **ODBC** — Direct database connection via TallyODBC driver
2. **XML API** — Tally's HTTP XML API on port 9000
3. **File Access** — Direct access to Tally data files
4. **Demo Mode** — Realistic sample data (no Tally required)

---

## Troubleshooting

### ODBC Connection Issues
- Ensure TallyPrime is running and ODBC driver is installed
- Verify DSN name in `tally-config.yml` matches Windows ODBC Data Sources
- Check Tally Gateway of Tally is enabled

### XML API Issues
- Ensure TallyPrime is running with port 9000 open
- Check firewall settings allow localhost:9000
- Verify Tally > F12 > Advanced Configuration > Enable ODBC Server

### Server Not Starting
- Check Node.js version: `node --version` (needs >= 18)
- Run `npm run build` before `npm start`
- Check for port conflicts

---

## Security Considerations

- The server communicates over stdio (no network exposure by default)
- ODBC credentials are stored in `tally-config.yml` — keep this file secure
- Never commit `tally-config.yml` with real credentials to version control
- Use environment variables for sensitive configuration in production
- Docker deployment uses a non-root user

---

## Performance Optimization

- Reports use in-memory generation for fast response times
- ODBC connection pooling configurable in `tally-config.yml`
- Cache TTL configurable (default: 5 minutes)
- Async operations throughout for non-blocking execution

---

## License

MIT © Emeren Smartsoft
