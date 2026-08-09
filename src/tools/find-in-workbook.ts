import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { WorkbookManager } from '../workbook/manager.js';
import { fail, ok } from './_helpers.js';
import { serializeCellValue } from '../util/serialize.js';
import { formatA1 } from '../util/cell-address.js';

const DEFAULT_MAX_HITS = 200;

export function registerFindInWorkbook(server: McpServer, manager: WorkbookManager): void {
  server.registerTool(
    'find_in_workbook',
    {
      title: 'Find in Workbook',
      description:
        'Search all sheets (or a subset) for cells whose evaluated value or formula contains the query. Supports plain substring or regex.',
      inputSchema: {
        workbookId: z.string(),
        query: z.string().min(1),
        regex: z.boolean().optional().describe('Interpret query as a regular expression'),
        caseSensitive: z.boolean().optional(),
        sheets: z.array(z.string()).optional().describe('Restrict to these sheet names'),
        searchFormulas: z.boolean().optional().describe('Also match formula strings (default true)'),
        maxHits: z.number().int().positive().max(5_000).optional(),
      },
    },
    async ({ workbookId, query, regex, caseSensitive, sheets, searchFormulas, maxHits }) => {
      try {
        const wb = manager.get(workbookId);
        const cap = maxHits ?? DEFAULT_MAX_HITS;
        const includeFormulas = searchFormulas !== false;

        const matcher = buildMatcher(query, { regex: !!regex, caseSensitive: !!caseSensitive });

        const targetSheets = sheets && sheets.length > 0 ? sheets : wb.metadata.sheets.map((s) => s.name);
        const hits: {
          sheet: string;
          cell: string;
          value: ReturnType<typeof serializeCellValue>;
          formula: string | null;
          matchedIn: 'value' | 'formula';
        }[] = [];

        outer: for (const sheetName of targetSheets) {
          const sheetMeta = wb.metadata.sheets.find(
            (s) => s.name === sheetName || s.name.toLowerCase() === sheetName.toLowerCase(),
          );
          if (!sheetMeta) continue;
          const sheetId = wb.hf.getSheetId(sheetMeta.name);
          if (sheetId === undefined) continue;
          const dims = wb.hf.getSheetDimensions(sheetId);
          for (let r = 0; r < dims.height; r++) {
            for (let c = 0; c < dims.width; c++) {
              const addr = { sheet: sheetId, row: r, col: c };
              const rawValue = wb.hf.getCellValue(addr);
              const value = serializeCellValue(rawValue);
              const valueStr = value === null ? '' : typeof value === 'object' ? value.error : String(value);
              let matchedIn: 'value' | 'formula' | null = null;
              if (matcher(valueStr)) matchedIn = 'value';
              let formula: string | null = null;
              if (includeFormulas) {
                const f = wb.hf.getCellFormula(addr);
                if (typeof f === 'string') {
                  formula = f;
                  if (!matchedIn && matcher(f)) matchedIn = 'formula';
                }
              }
              if (matchedIn) {
                hits.push({
                  sheet: sheetMeta.name,
                  cell: formatA1({ row: r, col: c }),
                  value,
                  formula,
                  matchedIn,
                });
                if (hits.length >= cap) break outer;
              }
            }
          }
        }

        return ok({
          workbookId,
          query,
          regex: !!regex,
          caseSensitive: !!caseSensitive,
          searchedSheets: targetSheets,
          hitCount: hits.length,
          truncated: hits.length >= cap,
          hits,
        });
      } catch (e) {
        return fail(e);
      }
    },
  );
}

function buildMatcher(
  query: string,
  opts: { regex: boolean; caseSensitive: boolean },
): (s: string) => boolean {
  if (opts.regex) {
    const re = new RegExp(query, opts.caseSensitive ? '' : 'i');
    return (s) => re.test(s);
  }
  if (opts.caseSensitive) {
    return (s) => s.includes(query);
  }
  const needle = query.toLowerCase();
  return (s) => s.toLowerCase().includes(needle);
}
