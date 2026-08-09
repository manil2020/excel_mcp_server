import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { WorkbookManager } from '../workbook/manager.js';
import { fail, ok } from './_helpers.js';

const NAMED_EXPR_PREVIEW = 50;

export function registerOpenWorkbook(server: McpServer, manager: WorkbookManager): void {
  server.registerTool(
    'open_workbook',
    {
      title: 'Open Excel Workbook',
      description:
        'Parse a local .xlsx/.xlsm file with ExcelJS and load it into HyperFormula. Returns a workbookId to pass to subsequent tools plus a summary of sheets and named expressions.',
      inputSchema: {
        path: z
          .string()
          .min(1)
          .describe('Absolute or relative path to a local .xlsx/.xlsm file. Environment variable EXCEL_MCP_ALLOWED_ROOTS can restrict allowed directories.'),
      },
    },
    async ({ path }) => {
      try {
        const wb = await manager.open(path);
        const named = wb.metadata.namedExpressions;
        return ok({
          workbookId: wb.id,
          filePath: wb.filePath,
          sizeBytes: wb.sizeBytes,
          loadedAt: wb.metadata.loadedAt,
          sheets: wb.metadata.sheets,
          namedExpressions: named.slice(0, NAMED_EXPR_PREVIEW),
          namedExpressionsTruncated: named.length > NAMED_EXPR_PREVIEW,
          hints: [
            'Pass workbookId as the first argument to list_sheets, read_range, get_cell, evaluate_formula, find_in_workbook, get_sheet_summary, get_workbook_stats.',
            'Call close_workbook when finished to release memory.',
          ],
        });
      } catch (e) {
        return fail(e);
      }
    },
  );
}
