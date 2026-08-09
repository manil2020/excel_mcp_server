import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { WorkbookManager, defaultManager } from './workbook/manager.js';
import { registerAllTools } from './tools/index.js';
import { log } from './util/logger.js';

export interface CreateServerOptions {
  manager?: WorkbookManager;
  name?: string;
  version?: string;
}

export interface CreatedServer {
  server: McpServer;
  manager: WorkbookManager;
}

export function createServer(opts: CreateServerOptions = {}): CreatedServer {
  const manager = opts.manager ?? defaultManager;
  const server = new McpServer(
    {
      name: opts.name ?? 'excel-mcp-server',
      version: opts.version ?? '0.1.0',
    },
    {
      instructions:
        'Explore Excel workbooks locally. Start by calling open_workbook with the file path, then use list_sheets, get_sheet_summary, read_range, get_cell, evaluate_formula, find_in_workbook, and get_workbook_stats. Values are evaluated by HyperFormula, so formulas resolve to computed values. Call close_workbook when finished.',
    },
  );
  registerAllTools(server, manager);
  log.info('MCP server configured', {
    name: opts.name ?? 'excel-mcp-server',
    version: opts.version ?? '0.1.0',
  });
  return { server, manager };
}
