import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { WorkbookManager } from '../workbook/manager.js';
import { fail, ok } from './_helpers.js';

export function registerListSheets(server: McpServer, manager: WorkbookManager): void {
  server.registerTool(
    'list_sheets',
    {
      title: 'List Sheets',
      description: 'List the sheets of an open workbook with dimensions.',
      inputSchema: {
        workbookId: z.string().describe('workbookId returned by open_workbook'),
      },
    },
    async ({ workbookId }) => {
      try {
        const wb = manager.get(workbookId);
        return ok({
          workbookId,
          filePath: wb.filePath,
          sheets: wb.metadata.sheets,
        });
      } catch (e) {
        return fail(e);
      }
    },
  );
}
