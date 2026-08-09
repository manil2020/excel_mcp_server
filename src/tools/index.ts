import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { WorkbookManager } from '../workbook/manager.js';
import { registerOpenWorkbook } from './open-workbook.js';
import { registerListWorkbooks } from './list-workbooks.js';
import { registerListSheets } from './list-sheets.js';
import { registerGetSheetSummary } from './get-sheet-summary.js';
import { registerReadRange } from './read-range.js';
import { registerGetCell } from './get-cell.js';
import { registerEvaluateFormula } from './evaluate-formula.js';
import { registerFindInWorkbook } from './find-in-workbook.js';
import { registerGetWorkbookStats } from './get-workbook-stats.js';
import { registerCloseWorkbook } from './close-workbook.js';

export function registerAllTools(server: McpServer, manager: WorkbookManager): void {
  registerOpenWorkbook(server, manager);
  registerListWorkbooks(server, manager);
  registerListSheets(server, manager);
  registerGetSheetSummary(server, manager);
  registerReadRange(server, manager);
  registerGetCell(server, manager);
  registerEvaluateFormula(server, manager);
  registerFindInWorkbook(server, manager);
  registerGetWorkbookStats(server, manager);
  registerCloseWorkbook(server, manager);
}

export const TOOL_NAMES = [
  'open_workbook',
  'list_workbooks',
  'list_sheets',
  'get_sheet_summary',
  'read_range',
  'get_cell',
  'evaluate_formula',
  'find_in_workbook',
  'get_workbook_stats',
  'close_workbook',
] as const;
export type ToolName = (typeof TOOL_NAMES)[number];
