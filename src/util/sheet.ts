import type { LoadedWorkbook } from '../workbook/types.js';
import { SheetNotFoundError } from './errors.js';

export function resolveSheet(wb: LoadedWorkbook, sheetName: string): { sheetId: number; name: string } {
  const available = wb.metadata.sheets.map((s) => s.name);
  const trimmed = sheetName.trim();
  const exact = available.find((n) => n === trimmed);
  const chosen = exact ?? available.find((n) => n.toLowerCase() === trimmed.toLowerCase());
  if (!chosen) throw new SheetNotFoundError(sheetName, available);
  const id = wb.hf.getSheetId(chosen);
  if (id === undefined) throw new SheetNotFoundError(sheetName, available);
  return { sheetId: id, name: chosen };
}
