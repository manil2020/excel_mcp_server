import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { WorkbookManager } from '../workbook/manager.js';
import { fail, ok } from './_helpers.js';

export function registerListWorkbooks(server: McpServer, manager: WorkbookManager): void {
  server.registerTool(
    'list_workbooks',
    {
      title: 'List Open Workbooks',
      description: 'List all workbooks currently loaded in this MCP server session, with paths and sheet counts.',
      inputSchema: {},
    },
    async () => {
      try {
        const items = manager.list().map((m) => ({
          workbookId: m.id,
          filePath: m.filePath,
          sizeBytes: m.sizeBytes,
          loadedAt: m.loadedAt,
          sheetCount: m.sheets.length,
          sheetNames: m.sheets.map((s) => s.name),
        }));
        return ok({
          count: items.length,
          limits: manager.limits,
          workbooks: items,
        });
      } catch (e) {
        return fail(e);
      }
    },
  );
}
