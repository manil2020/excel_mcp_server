# Testing the excel-mcp-server locally

Step-by-step guide for hooking the server up to VS Code or Claude Desktop and driving it against the sample workbooks in `samples/`.

## 1. Prerequisites

```bash
cd <path-to-repo>/excel_mcp_server

npm install
npm run build          # writes dist/index.js (executable)
npm test               # 10/10 should pass
```

Generate the sample workbooks (only needed once):

```bash
npx tsx scripts/generate-samples.ts
```

That produces four synthetic files in `samples/`. There is also `samples/financial-sample.xlsx` — Microsoft's public Power BI Financial Sample — downloaded separately.

Verify all five workbooks open via the server:

```bash
node --input-type=module -e "
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { readdirSync } from 'node:fs'; import { resolve } from 'node:path';
const c = new Client({ name: 'smoke', version: '0.0.0' });
await c.connect(new StdioClientTransport({ command: 'node', args: ['dist/index.js'] }));
for (const f of readdirSync('samples').filter(x=>x.endsWith('.xlsx')).sort()) {
  const r = await c.callTool({ name: 'open_workbook', arguments: { path: resolve('samples', f) } });
  console.log(f, '=>', r.isError ? 'FAIL' : 'OK');
  if (!r.isError) await c.callTool({ name: 'close_workbook', arguments: { workbookId: r.structuredContent.workbookId } });
}
await c.close();
"
```

## 2. Sample workbook cheat-sheet

| File | Sheets | Highlights | Good for testing |
|------|--------|------------|------------------|
| `financial-sample.xlsx` | `Sheet1` (700 rows × 16 cols) | Real Microsoft Power BI sample — country, segment, product, sales, profit, discount | `get_sheet_summary`, `find_in_workbook`, `evaluate_formula` (e.g. `=SUMIFS(...)`) |
| `finance-quarterly.xlsx` | Q1..Q4 + `Yearly` | Cross-sheet SUMIFS, named range `FullYearRevenue`, quarterly totals feeding a yearly rollup | multi-sheet references, `list_sheets`, `read_range`, named expressions |
| `employee-data.xlsx` | `Employees`, `SalaryBands`, `Summary` | `DATEDIF`, nested `IF`, `COUNTIF`/`SUMIF` across sheets, real dates | `get_cell` with formulas, date handling |
| `messy-report.xlsx` | `Monthly Report`, `Notes` | Merged title cell, headers on row 5 (not row 1), currency formatting, notes sheet | header-detection heuristic in `get_sheet_summary`, `find_in_workbook` on free text |
| `transactions-2k.xlsx` | `Transactions` (2000 rows × 8 cols) | Dates, quantities, computed totals | `read_range` with `maxCells`, `get_workbook_stats`, `find_in_workbook` performance |

## 3. Wire the server into VS Code (Copilot Chat MCP)

Add this to `.vscode/mcp.json` in whatever workspace you want to chat from (or user `mcp.json`):

```json
{
  "servers": {
    "excel": {
      "command": "node",
      "args": ["<path-to-repo>/excel_mcp_server/dist/index.js"],
      "env": {
        "EXCEL_MCP_LOG_LEVEL": "info",
        "EXCEL_MCP_ALLOWED_ROOTS": "<path-to-repo>/excel_mcp_server/samples:~/Desktop"
      }
    }
  }
}
```

Setting `EXCEL_MCP_ALLOWED_ROOTS` restricts which directories the server will open — safer default than "any local path".

Reload VS Code, open Copilot Chat, and switch to Agent mode. The `excel` tools should appear.

## 4. Wire the server into Claude Desktop

Edit `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": ["<path-to-repo>/excel_mcp_server/dist/index.js"],
      "env": {
        "EXCEL_MCP_LOG_LEVEL": "info",
        "EXCEL_MCP_ALLOWED_ROOTS": "<path-to-repo>/excel_mcp_server/samples"
      }
    }
  }
}
```

Quit and restart Claude Desktop. The tools appear under the hammer icon.

## 5. Prompts to try

Paste these into your MCP client to exercise different features.

### Basics
> Open `<path-to-repo>/excel_mcp_server/samples/financial-sample.xlsx`, list the sheets, and give me a summary of the first sheet including headers and column types.

### Cross-sheet formulas
> Open `samples/finance-quarterly.xlsx`. What was the full-year revenue, cost, and profit margin? Which quarter had the highest margin, and by how much?

### Search & drill-down
> In `samples/financial-sample.xlsx`, find every row where the country is "Germany" and Segment is "Enterprise". What was the total profit for that combination?

### Ad-hoc formula evaluation
> Using `samples/financial-sample.xlsx`, compute `=SUMIFS(Sheet1!'Gross Sales', Sheet1!Country, "United States", Sheet1!Segment, "Government")`. Then also give me the count of rows matching those criteria.

### Nested formulas & dates
> Open `samples/employee-data.xlsx`. For each department in the Summary sheet, list headcount, total payroll, and average salary. Then identify the top three employees by tenure years.

### Messy-workbook resilience
> Open `samples/messy-report.xlsx`. Where is the actual data table? Extract the SKUs and their line totals as a clean JSON array.

### Large-range performance
> Open `samples/transactions-2k.xlsx`. Read the first 20 rows, then compute total revenue per country and per product using `evaluate_formula`.

### Workbook profiling
> Give me `get_workbook_stats` on every sample workbook in `samples/`. Rank them by number of formula cells.

## 6. Troubleshooting

**"WORKBOOK_NOT_FOUND"** — the workbookId doesn't exist in this session (you closed it, or the server restarted). Call `open_workbook` again.

**"FILE_ACCESS: path is outside the allowed roots"** — the requested file is not inside any directory listed in `EXCEL_MCP_ALLOWED_ROOTS`. Add the parent directory to that env var, or remove the restriction entirely.

**"FILE_TOO_LARGE"** — the workbook is over `EXCEL_MCP_MAX_FILE_MB` (default 100). Raise the cap: `"EXCEL_MCP_MAX_FILE_MB": "500"`.

**Client hangs on startup** — verify the built binary runs manually:
```bash
node <path-to-repo>/excel_mcp_server/dist/index.js
# then in another terminal, kill it: Ctrl+C
```
If nothing prints to stderr, `npm run build` has not been run.

**Tools don't appear in the client** — the MCP config path is off, or the client hasn't been fully restarted. In VS Code, run `MCP: List Servers` from the command palette to see status. In Claude Desktop, fully quit (Cmd+Q) and reopen.

**Formula returns a `{ error, type }` object** — HyperFormula raised an Excel-style error (`#DIV/0!`, `#REF!`, etc.). The `type` field tells you which one. Fix the referenced cells or the formula.

## 7. Useful logs

Server logs go to **stderr** as JSON lines. To inspect what's happening, wrap the command in a shell that captures stderr, or set `"EXCEL_MCP_LOG_LEVEL": "debug"` in the client env for verbose output. Stdout is reserved for MCP JSON-RPC framing and must not be polluted.
