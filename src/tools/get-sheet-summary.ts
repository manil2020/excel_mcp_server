import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { WorkbookManager } from '../workbook/manager.js';
import { fail, ok } from './_helpers.js';
import { resolveSheet } from '../util/sheet.js';
import { serializeCellValue } from '../util/serialize.js';
import { formatA1 } from '../util/cell-address.js';

export function registerGetSheetSummary(server: McpServer, manager: WorkbookManager): void {
  server.registerTool(
    'get_sheet_summary',
    {
      title: 'Get Sheet Summary',
      description:
        'Return dimensions, a heuristic header row, per-column type inference, and a sample of the first rows for a sheet.',
      inputSchema: {
        workbookId: z.string(),
        sheet: z.string().describe('Sheet name (case-insensitive)'),
        sampleRows: z.number().int().min(1).max(500).optional().describe('Rows to include in the sample (default 10)'),
      },
    },
    async ({ workbookId, sheet, sampleRows }) => {
      try {
        const wb = manager.get(workbookId);
        const { sheetId, name } = resolveSheet(wb, sheet);
        const dims = wb.hf.getSheetDimensions(sheetId);
        const rowCount = dims.height;
        const colCount = dims.width;
        const rowsToRead = Math.min(sampleRows ?? 10, rowCount);

        const values: (ReturnType<typeof serializeCellValue>)[][] = [];
        for (let r = 0; r < rowsToRead; r++) {
          const row: (ReturnType<typeof serializeCellValue>)[] = [];
          for (let c = 0; c < colCount; c++) {
            row.push(serializeCellValue(wb.hf.getCellValue({ sheet: sheetId, row: r, col: c })));
          }
          values.push(row);
        }

        const headers = detectHeaderRow(values, colCount);
        const columnTypes = inferColumnTypes(values, colCount, headers ? 1 : 0);

        return ok({
          workbookId,
          sheet: name,
          rowCount,
          colCount,
          topLeft: formatA1({ row: 0, col: 0 }),
          bottomRight: rowCount > 0 && colCount > 0 ? formatA1({ row: rowCount - 1, col: colCount - 1 }) : null,
          headers,
          columnTypes,
          sampleRowCount: values.length,
          sample: values,
        });
      } catch (e) {
        return fail(e);
      }
    },
  );
}

function detectHeaderRow(rows: unknown[][], colCount: number): string[] | null {
  if (rows.length === 0 || colCount === 0) return null;
  const first = rows[0] ?? [];
  let stringCount = 0;
  let nonNullCount = 0;
  for (let c = 0; c < colCount; c++) {
    const v = first[c];
    if (v === null || v === undefined || v === '') continue;
    nonNullCount++;
    if (typeof v === 'string') stringCount++;
  }
  if (nonNullCount === 0) return null;
  if (stringCount / nonNullCount >= 0.7) {
    return first.slice(0, colCount).map((v) => (v === null || v === undefined ? '' : String(v)));
  }
  return null;
}

type ColType = 'number' | 'string' | 'boolean' | 'date' | 'mixed' | 'empty';

function inferColumnTypes(rows: unknown[][], colCount: number, startRow: number): ColType[] {
  const out: ColType[] = [];
  for (let c = 0; c < colCount; c++) {
    const seen = new Set<string>();
    for (let r = startRow; r < rows.length; r++) {
      const v = rows[r]?.[c];
      if (v === null || v === undefined || v === '') continue;
      if (typeof v === 'number') seen.add('number');
      else if (typeof v === 'boolean') seen.add('boolean');
      else if (typeof v === 'string') seen.add('string');
      else seen.add('string');
    }
    if (seen.size === 0) out.push('empty');
    else if (seen.size === 1) out.push([...seen][0] as ColType);
    else out.push('mixed');
  }
  return out;
}
