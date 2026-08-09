import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { WorkbookManager } from '../workbook/manager.js';
import { fail, ok } from './_helpers.js';

export function registerGetWorkbookStats(server: McpServer, manager: WorkbookManager): void {
  server.registerTool(
    'get_workbook_stats',
    {
      title: 'Get Workbook Stats',
      description: 'Return aggregate statistics for an open workbook (per-sheet cell counts, formula counts, errors).',
      inputSchema: {
        workbookId: z.string(),
      },
    },
    async ({ workbookId }) => {
      try {
        const wb = manager.get(workbookId);
        const perSheet: {
          sheet: string;
          rows: number;
          cols: number;
          nonEmptyCells: number;
          formulaCells: number;
          errorCells: number;
        }[] = [];
        let totalNonEmpty = 0;
        let totalFormulas = 0;
        let totalErrors = 0;

        for (const s of wb.metadata.sheets) {
          const sheetId = wb.hf.getSheetId(s.name);
          if (sheetId === undefined) continue;
          const dims = wb.hf.getSheetDimensions(sheetId);
          let nonEmpty = 0;
          let formulas = 0;
          let errors = 0;
          for (let r = 0; r < dims.height; r++) {
            for (let c = 0; c < dims.width; c++) {
              const addr = { sheet: sheetId, row: r, col: c };
              const isEmpty = wb.hf.isCellEmpty(addr);
              if (!isEmpty) nonEmpty++;
              if (wb.hf.doesCellHaveFormula(addr)) formulas++;
              const v = wb.hf.getCellValue(addr);
              if (v && typeof v === 'object' && 'type' in v && 'value' in v) errors++;
            }
          }
          perSheet.push({
            sheet: s.name,
            rows: dims.height,
            cols: dims.width,
            nonEmptyCells: nonEmpty,
            formulaCells: formulas,
            errorCells: errors,
          });
          totalNonEmpty += nonEmpty;
          totalFormulas += formulas;
          totalErrors += errors;
        }

        return ok({
          workbookId,
          filePath: wb.filePath,
          sizeBytes: wb.sizeBytes,
          loadedAt: wb.metadata.loadedAt,
          sheetCount: wb.metadata.sheets.length,
          namedExpressionCount: wb.metadata.namedExpressions.length,
          totals: {
            nonEmptyCells: totalNonEmpty,
            formulaCells: totalFormulas,
            errorCells: totalErrors,
          },
          perSheet,
        });
      } catch (e) {
        return fail(e);
      }
    },
  );
}
