import type ExcelJS from 'exceljs';
import type { HyperFormula } from 'hyperformula';

/** Value shape accepted by HyperFormula for a single cell. */
export type HfCellValue = string | number | boolean | null;

export interface SheetMetadata {
  name: string;
  index: number; // 0-based sheet index
  rowCount: number;
  colCount: number;
}

export interface WorkbookMetadata {
  id: string;
  filePath: string;
  sizeBytes: number;
  loadedAt: string; // ISO
  sheets: SheetMetadata[];
  namedExpressions: { name: string; expression: string; scope: string | null }[];
}

export interface LoadedWorkbook {
  id: string;
  filePath: string;
  sizeBytes: number;
  loadedAt: Date;
  hf: HyperFormula;
  excel: ExcelJS.Workbook;
  metadata: WorkbookMetadata;
}
