# excel-mcp-server

Local **Model Context Protocol (MCP)** server that lets LLMs *explore* Excel workbooks. It combines:

- **[ExcelJS](https://github.com/exceljs/exceljs)** — parses `.xlsx` / `.xlsm` files (values, formulas, rich text, hyperlinks, dates).
- **[HyperFormula](https://github.com/handsontable/hyperformula)** — an in-memory spreadsheet engine that re-evaluates every formula, so tools return the *computed* result rather than the value cached in the file.

The server speaks MCP over **stdio**, so it works with any MCP-compatible client: **VS Code**, **Claude Desktop**, **Cursor**, and custom agents.

> **License**: [GPL-3.0-or-later](LICENSE). HyperFormula is dual-licensed (GPL-3.0 or commercial); see [NOTICE](NOTICE) for details.

---

## Intended use

> **This repository is a personal, educational proof-of-concept.** Please read this section before reusing the code.

**Do not:**

- Bundle this project inside a commercial software product.
- Deploy this project as part of a paid SaaS offering.
- Redistribute this project (or its binaries) under any license other than GPL-3.0-or-later.
- Use this project as a substitute for a properly licensed HyperFormula commercial integration.

**Why the caveat?** HyperFormula (see the [Credits](#credits--acknowledgements) section below) is dual-licensed. This project uses the open-source GPL-3.0 tier by passing `licenseKey: 'gpl-v3'` at engine construction. If you plan to embed HyperFormula in commercial or closed-source software, you **must** obtain a commercial license directly from [Handsontable](https://handsontable.com/pricing) and use your commercial license key. Doing so is a requirement of HyperFormula's dual-licensing model, not a limitation of this repository.

Everything else in this repository (ExcelJS, MCP SDK, Zod, TypeScript, etc.) is MIT / Apache-2.0 licensed and free to reuse under those terms.

If your use case is purely personal exploration, learning, or open-source contribution back to a GPL-3.0-or-later project, you are welcome to use, fork, and modify this code freely.

---

## How this was built

This repository was scaffolded and iterated on with the help of an AI coding assistant (Anthropic's Claude, via GitHub Copilot's agent mode in VS Code). The design decisions, dependency choices, licensing posture, and architectural trade-offs were driven by the author; the assistant contributed code generation, test scaffolding, and documentation drafting under human review. All committed code has been read and vetted by a human before landing.

---

## Features

The server exposes 10 read-only tools that cover the common "explore this spreadsheet" workflow:

| Tool                  | Purpose                                                                 |
| --------------------- | ----------------------------------------------------------------------- |
| `open_workbook`       | Load a local `.xlsx`/`.xlsm` file and get back a `workbookId`.          |
| `list_workbooks`      | List all workbooks currently loaded in the session.                     |
| `list_sheets`         | Enumerate sheets with row/column counts.                                |
| `get_sheet_summary`   | Dimensions, header row heuristic, per-column type inference, sample.    |
| `read_range`          | Read a rectangular A1 range with computed values and optional formulas. |
| `get_cell`            | Value + formula + cell type + number format for a single cell.          |
| `evaluate_formula`    | Evaluate an arbitrary Excel formula against the loaded workbook.        |
| `find_in_workbook`    | Substring / regex search across all sheets (values and formulas).       |
| `get_workbook_stats`  | Aggregate counts (non-empty, formulas, errors) per sheet and overall.   |
| `close_workbook`      | Release memory held by a workbook.                                      |

Highlights:

- Formulas are re-evaluated by HyperFormula — you always see the *live* computed value, never a stale cached one.
- Handles cross-sheet references (`=SUM(Data!C2:C5)`), named expressions defined in the file, and standard Excel error types (`#DIV/0!`, `#N/A`, …).
- Safe by default: read-only, size- and cell-count-capped, and an optional allow-list restricts which directories can be opened.
- No writes, no network calls, no telemetry.

---

## Install & run

Prerequisites: **Node.js ≥ 18.17** (Node 20+ recommended).

```bash
git clone https://github.com/manil2020/excel_mcp_server.git
cd excel_mcp_server
npm install
npm run build
npm test        # runs the vitest suite

# Start the server on stdio (this is what MCP clients do automatically)
node dist/index.js
```

For local development without a build step:

```bash
npm run dev
```

---

## Wiring the server into MCP clients

### VS Code (GitHub Copilot Chat MCP)

Add the server to your workspace or user MCP config. A minimal `.vscode/mcp.json`:

```json
{
  "servers": {
    "excel": {
      "command": "node",
      "args": ["/absolute/path/to/excel_mcp_server/dist/index.js"],
      "env": {
        "EXCEL_MCP_LOG_LEVEL": "info"
      }
    }
  }
}
```

See [examples/vscode-mcp.json](examples/vscode-mcp.json) for a ready-to-copy version.

### Claude Desktop

Edit `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS) or the equivalent on Windows / Linux:

```json
{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": ["/absolute/path/to/excel_mcp_server/dist/index.js"]
    }
  }
}
```

See [examples/claude-desktop-config.json](examples/claude-desktop-config.json).

### Any other MCP client

Point the client at `node /absolute/path/to/dist/index.js` with the environment variables described below. The server uses stdio, the standard MCP local transport.

---

## Configuration

All configuration is via environment variables. Defaults are safe for a laptop workload.

| Variable                    | Default | Description                                                                                        |
| --------------------------- | ------- | -------------------------------------------------------------------------------------------------- |
| `EXCEL_MCP_MAX_FILE_MB`     | `100`   | Reject workbooks larger than this many megabytes.                                                  |
| `EXCEL_MCP_MAX_WORKBOOKS`   | `8`     | Maximum number of workbooks kept in memory concurrently.                                           |
| `EXCEL_MCP_ALLOWED_ROOTS`   | *(unset)* | Colon-separated absolute paths. If set, only files under one of these roots may be opened.       |
| `EXCEL_MCP_LOG_LEVEL`       | `info`  | `debug` / `info` / `warn` / `error`. Logs go to stderr as JSON lines.                              |

---

## Example workflow (LLM-side)

```text
open_workbook(path="/data/orders.xlsx")
→ { workbookId: "wb_a1b2c3d4", sheets: [ ... ] }

get_sheet_summary(workbookId, sheet="Orders")
→ headers: ["order_id", "customer", "amount", "region"], columnTypes: [ ... ]

find_in_workbook(workbookId, query="EMEA")
→ hits: [ { sheet: "Orders", cell: "D42", value: "EMEA" }, ... ]

evaluate_formula(workbookId, formula="=SUMIFS(Orders!C:C, Orders!D:D, \"EMEA\")")
→ { result: 12345.67 }

close_workbook(workbookId)
```

---

## Architecture

- **`src/index.ts`** – Bin entry: creates the server and connects a `StdioServerTransport`.
- **`src/server.ts`** – Constructs `McpServer`, wires all tools, exports `createServer()` for tests.
- **`src/workbook/loader.ts`** – Reads the file with ExcelJS, converts each cell into HyperFormula-friendly primitives (formulas kept as `=…` strings, dates converted to Excel serial numbers, rich text flattened), then calls `HyperFormula.buildFromSheets`.
- **`src/workbook/manager.ts`** – Registry of open workbooks keyed by opaque `wb_*` handles, with per-process limits and graceful shutdown.
- **`src/tools/*.ts`** – One file per MCP tool. Each registers itself with the shared `McpServer` and delegates to the workbook manager.

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for a diagram and the full data flow.

---

## Development

```bash
npm run dev          # tsx watch (single run)
npm run typecheck    # tsc --noEmit
npm run lint         # eslint
npm run format       # prettier --write
npm test             # vitest run
npm run test:watch   # vitest watch
```

Directory layout:

```
excel_mcp_server/
├── src/
│   ├── index.ts              # #!/usr/bin/env node entry
│   ├── server.ts             # MCP server factory
│   ├── workbook/             # ExcelJS + HyperFormula bridge
│   ├── tools/                # One file per MCP tool
│   └── util/                 # Address helpers, errors, logger, path safety
├── tests/
│   ├── fixtures/             # Programmatic xlsx generation
│   ├── loader.test.ts        # Workbook layer
│   └── tools.test.ts         # Full MCP round-trip (in-memory transport)
├── docs/
│   ├── ARCHITECTURE.md
│   └── USAGE.md
├── examples/
│   ├── vscode-mcp.json
│   └── claude-desktop-config.json
├── LICENSE                   # GPL-3.0
└── NOTICE                    # Third-party licensing
```

---

## Licensing summary

This project uses [HyperFormula](https://github.com/handsontable/hyperformula), which is dual-licensed under GPL-3.0 or a commercial license from Handsontable. Because this project links against HyperFormula, the combined work is distributed as **GPL-3.0-or-later**. See [NOTICE](NOTICE) for full detail and instructions for commercial reuse.

The other runtime dependencies (ExcelJS, `@modelcontextprotocol/sdk`, zod) are MIT-licensed and compatible with GPL-3.0.

---

## Credits & acknowledgements

This project would not exist without the following open-source libraries and the maintainers who built them. If you find this project useful, please star and support the upstream projects — they did the hard work.

### Runtime dependencies

| Project | Maintainer(s) | License | Homepage |
|---------|---------------|---------|----------|
| **HyperFormula** | Handsontable Sp. z o.o. (Poland) | GPL-3.0 / Commercial | [hyperformula.handsontable.com](https://hyperformula.handsontable.com/) · [source](https://github.com/handsontable/hyperformula) |
| **ExcelJS** | Guyon Roche and contributors | MIT | [github.com/exceljs/exceljs](https://github.com/exceljs/exceljs) |
| **@modelcontextprotocol/sdk** | Anthropic PBC and the MCP community | MIT | [github.com/modelcontextprotocol/typescript-sdk](https://github.com/modelcontextprotocol/typescript-sdk) |
| **zod** | Colin McDonnell ([@colinhacks](https://github.com/colinhacks)) and contributors | MIT | [zod.dev](https://zod.dev/) |

### Development dependencies

| Project | Maintainer(s) | License |
|---------|---------------|---------|
| **TypeScript** | Microsoft | Apache-2.0 |
| **Vitest** | Anthony Fu ([@antfu](https://github.com/antfu)) and contributors | MIT |
| **tsx** | Hiroki Osame ([@privatenumber](https://github.com/privatenumber)) | MIT |
| **ESLint** | OpenJS Foundation / ESLint team | MIT |
| **Prettier** | Prettier team | MIT |
| **Node.js** | OpenJS Foundation | MIT |

### Protocol

The **Model Context Protocol** specification is developed and stewarded by [Anthropic](https://www.anthropic.com/) and the wider MCP community at [modelcontextprotocol.io](https://modelcontextprotocol.io/). The protocol design work makes servers like this one possible.

### Sample data

- The `samples/financial-sample.xlsx` fixture (used for local testing only, gitignored) is **Microsoft's public Power BI Financial Sample workbook**, freely distributed by Microsoft for learning purposes: [download link](https://go.microsoft.com/fwlink/?LinkID=521962).
- All other files in `samples/` are synthesised locally by [scripts/generate-samples.ts](scripts/generate-samples.ts) and contain no real personal or commercial data.

If any maintainer name or attribution detail here is missing or wrong, please open an issue — corrections are welcome.

---

## Roadmap

- Optional write tools (`set_cell`, `save_workbook`) behind an opt-in `--allow-writes` flag.
- Streaming for very large ranges via `read_range` pagination.
- CSV / `.xls` legacy format support.
- Prompt / resource providers exposing the workbook as MCP `Resources`.
