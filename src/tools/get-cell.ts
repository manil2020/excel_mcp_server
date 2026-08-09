import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { WorkbookManager } from '../workbook/manager.js';
import { fail, ok } from './_helpers.js';
import { resolveSheet } from '../util/sheet.js';
import { serializeCellValue } from '../util/serialize.js';
import { parseA1 } from '../util/cell-address.js';

export function registerGetCell(server: McpServer, manager: WorkbookManager): void {
  server.registerTool(
    'get_cell',
    {
      title: 'Get Cell',
      description: 'Return the evaluated value, formula (if any), and cell type for a single cell.',
      inputSchema: {
        workbookId: z.string(),
        sheet: z.string(),
        cell: z.string().describe('A1 address like "B7"'),
      },
    },
    async ({ workbookId, sheet, cell }) => {
      try {
        const wb = manager.get(workbookId);
        const { sheetId, name } = resolveSheet(wb, sheet);
        const coord = parseA1(cell);
        const address = { sheet: sheetId, row: coord.row, col: coord.col };
        const value = serializeCellValue(wb.hf.getCellValue(address));
        const formula = wb.hf.getCellFormula(address);
        const cellType = wb.hf.getCellType(address);
        const valueType = wb.hf.getCellValueType(address);

        // Read raw ExcelJS value for extra context (formatting, hyperlink text, etc.).
        const excelSheet = wb.excel.getWorksheet(name);
        const excelCell = excelSheet ? excelSheet.getCell(coord.row + 1, coord.col + 1) : undefined;
        const rawExcelValue = excelCell?.value ?? null;
        const numFmt = excelCell?.numFmt ?? null;

        return ok({
          workbookId,
          sheet: name,
          cell,
          value,
          formula: typeof formula === 'string' ? formula : null,
          cellType,
          valueType,
          numberFormat: numFmt,
          excelRaw: coerceRawForJson(rawExcelValue),
        });
      } catch (e) {
        return fail(e);
      }
    },
  );
}

function coerceRawForJson(v: unknown): unknown {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'object') {
    const obj = v as Record<string, unknown>;
    if (Array.isArray(obj.richText)) return (obj.richText as { text?: string }[]).map((r) => r.text ?? '').join('');
    if ('formula' in obj || 'sharedFormula' in obj) {
      return {
        formula: obj.formula ?? obj.sharedFormula ?? null,
        result: obj.result ?? null,
      };
    }
    if ('text' in obj && 'hyperlink' in obj) {
      return { text: obj.text, hyperlink: obj.hyperlink };
    }
    if ('error' in obj) return { error: obj.error };
  }
  return v;
}
