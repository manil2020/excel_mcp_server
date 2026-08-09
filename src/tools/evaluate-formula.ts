import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { WorkbookManager } from '../workbook/manager.js';
import { fail, ok } from './_helpers.js';
import { resolveSheet } from '../util/sheet.js';
import { serializeCellValue } from '../util/serialize.js';

export function registerEvaluateFormula(server: McpServer, manager: WorkbookManager): void {
  server.registerTool(
    'evaluate_formula',
    {
      title: 'Evaluate Formula',
      description:
        'Evaluate an arbitrary Excel-style formula in the context of an open workbook using HyperFormula. The formula may reference any sheet by name (e.g. `=SUM(Sheet1!A1:A20)`).',
      inputSchema: {
        workbookId: z.string(),
        formula: z.string().describe('Formula to evaluate. A leading "=" is added automatically if missing.'),
        sheet: z
          .string()
          .optional()
          .describe('Sheet name that provides the context for relative references (defaults to the first sheet).'),
      },
    },
    async ({ workbookId, formula, sheet }) => {
      try {
        const wb = manager.get(workbookId);
        const contextName = sheet ?? wb.metadata.sheets[0]?.name;
        if (!contextName) throw new Error('workbook has no sheets to provide a context for evaluation');
        const { sheetId, name } = resolveSheet(wb, contextName);
        const expr = formula.trimStart().startsWith('=') ? formula.trimStart() : `=${formula.trimStart()}`;
        const result = wb.hf.calculateFormula(expr, sheetId);
        const serialized = Array.isArray(result)
          ? (result as unknown[][]).map((row) => row.map((v) => serializeCellValue(v)))
          : serializeCellValue(result);
        return ok({
          workbookId,
          contextSheet: name,
          formula: expr,
          resultType: Array.isArray(result) ? 'range' : 'scalar',
          result: serialized,
        });
      } catch (e) {
        return fail(e);
      }
    },
  );
}
