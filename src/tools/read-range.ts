import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { WorkbookManager } from '../workbook/manager.js';
import { fail, ok } from './_helpers.js';
import { resolveSheet } from '../util/sheet.js';
import { serializeCellValue } from '../util/serialize.js';
import { formatA1, parseRange, rangeSize } from '../util/cell-address.js';
import { InvalidRangeError } from '../util/errors.js';

const DEFAULT_MAX_CELLS = 5_000;

export function registerReadRange(server: McpServer, manager: WorkbookManager): void {
  server.registerTool(
    'read_range',
    {
      title: 'Read Range',
      description:
        'Read a rectangular range from a sheet. Returns evaluated values (via HyperFormula) and, optionally, the corresponding formulas.',
      inputSchema: {
        workbookId: z.string(),
        sheet: z.string(),
        range: z.string().describe('A1 range like "A1:D50" or a single cell like "B2"'),
        includeFormulas: z.boolean().optional().describe('If true, also return the formula string per cell where present'),
        maxCells: z
          .number()
          .int()
          .positive()
          .max(100_000)
          .optional()
          .describe(`Safety cap on total cells read (default ${DEFAULT_MAX_CELLS})`),
      },
    },
    async ({ workbookId, sheet, range, includeFormulas, maxCells }) => {
      try {
        const wb = manager.get(workbookId);
        const { sheetId, name } = resolveSheet(wb, sheet);
        const parsed = parseRange(range);
        const size = rangeSize(parsed);
        const cap = maxCells ?? DEFAULT_MAX_CELLS;
        if (size.cells > cap) {
          throw new InvalidRangeError(
            range,
            `range has ${size.cells} cells which exceeds maxCells=${cap}. Narrow the range or raise maxCells.`,
          );
        }
        const values: ReturnType<typeof serializeCellValue>[][] = [];
        const formulas: (string | null)[][] | undefined = includeFormulas ? [] : undefined;
        for (let r = parsed.start.row; r <= parsed.end.row; r++) {
          const rowVals: ReturnType<typeof serializeCellValue>[] = [];
          const rowFormulas: (string | null)[] | undefined = includeFormulas ? [] : undefined;
          for (let c = parsed.start.col; c <= parsed.end.col; c++) {
            rowVals.push(serializeCellValue(wb.hf.getCellValue({ sheet: sheetId, row: r, col: c })));
            if (rowFormulas) {
              const f = wb.hf.getCellFormula({ sheet: sheetId, row: r, col: c });
              rowFormulas.push(typeof f === 'string' ? f : null);
            }
          }
          values.push(rowVals);
          if (rowFormulas && formulas) formulas.push(rowFormulas);
        }

        return ok({
          workbookId,
          sheet: name,
          range: `${formatA1(parsed.start)}:${formatA1(parsed.end)}`,
          rows: size.rows,
          cols: size.cols,
          values,
          ...(formulas ? { formulas } : {}),
        });
      } catch (e) {
        return fail(e);
      }
    },
  );
}
