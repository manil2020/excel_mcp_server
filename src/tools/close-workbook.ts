import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { WorkbookManager } from '../workbook/manager.js';
import { fail, ok } from './_helpers.js';

export function registerCloseWorkbook(server: McpServer, manager: WorkbookManager): void {
  server.registerTool(
    'close_workbook',
    {
      title: 'Close Workbook',
      description: 'Release the resources held by an open workbook.',
      inputSchema: {
        workbookId: z.string(),
      },
    },
    async ({ workbookId }) => {
      try {
        const closed = manager.close(workbookId);
        return ok({ workbookId, closed });
      } catch (e) {
        return fail(e);
      }
    },
  );
}
