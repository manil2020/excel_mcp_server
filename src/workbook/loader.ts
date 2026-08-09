import ExcelJS from 'exceljs';
import { HyperFormula } from 'hyperformula';
import * as fs from 'node:fs';
import { assertReadableFile, resolveSafePath } from '../util/paths.js';
import { FileTooLargeError, UnsupportedFileError } from '../util/errors.js';
import { log } from '../util/logger.js';
import type { HfCellValue, LoadedWorkbook, SheetMetadata, WorkbookMetadata } from './types.js';

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MAX_CELLS_PER_SHEET = 5_000_000;

function jsDateToExcelSerial(d: Date): number {
  return (d.getTime() - EXCEL_EPOCH_MS) / 86_400_000;
}

function primitiveOrNull(v: unknown): HfCellValue {
  if (v === null || v === undefined) return null;
  const t = typeof v;
  if (t === 'string' || t === 'number' || t === 'boolean') return v as HfCellValue;
  return String(v);
}

/**
 * Convert an ExcelJS cell value into the primitive shape HyperFormula expects.
 * Formulas are emitted as "=..." strings so HF re-evaluates them, giving us a
 * single source of truth for computed values.
 */
function cellToHfValue(cell: ExcelJS.Cell): HfCellValue {
  const value = cell.value;
  if (value === null || value === undefined) return null;

  if (typeof value === 'object') {
    const obj = value as unknown as Record<string, unknown>;

    if (typeof obj.formula === 'string' && obj.formula.length > 0) {
      return `=${obj.formula}`;
    }
    if (typeof obj.sharedFormula === 'string' && obj.sharedFormula.length > 0) {
      return `=${obj.sharedFormula}`;
    }
    if (Array.isArray(obj.richText)) {
      return (obj.richText as { text?: string }[]).map((r) => r.text ?? '').join('');
    }
    if ('text' in obj) {
      return primitiveOrNull(obj.text);
    }
    if ('error' in obj) {
      return String(obj.error);
    }
    if (value instanceof Date) {
      return jsDateToExcelSerial(value);
    }
    if ('result' in obj) {
      return primitiveOrNull(obj.result);
    }
  }

  return primitiveOrNull(value);
}

function extractSheetGrid(worksheet: ExcelJS.Worksheet): HfCellValue[][] {
  const rowMax = worksheet.rowCount ?? 0;
  const colMax = worksheet.columnCount ?? 0;
  if (rowMax === 0 || colMax === 0) return [];
  if (rowMax * colMax > MAX_CELLS_PER_SHEET) {
    throw new UnsupportedFileError(
      worksheet.name,
      `sheet "${worksheet.name}" has ${rowMax}x${colMax} cells which exceeds the safety cap of ${MAX_CELLS_PER_SHEET.toLocaleString()} cells per sheet`,
    );
  }
  const grid: HfCellValue[][] = new Array(rowMax);
  for (let r = 1; r <= rowMax; r++) {
    const row: HfCellValue[] = new Array(colMax);
    const excelRow = worksheet.getRow(r);
    for (let c = 1; c <= colMax; c++) {
      row[c - 1] = cellToHfValue(excelRow.getCell(c));
    }
    grid[r - 1] = row;
  }
  return grid;
}

export interface LoadOptions {
  id: string;
  maxFileBytes: number;
}

export async function loadWorkbook(inputPath: string, opts: LoadOptions): Promise<LoadedWorkbook> {
  const absPath = resolveSafePath(inputPath);
  const stat = assertReadableFile(absPath);
  if (stat.size > opts.maxFileBytes) {
    throw new FileTooLargeError(absPath, stat.size, opts.maxFileBytes);
  }
  const ext = absPath.toLowerCase().split('.').pop() ?? '';
  if (!['xlsx', 'xlsm'].includes(ext)) {
    throw new UnsupportedFileError(absPath, `extension ".${ext}" is not supported (accepted: .xlsx, .xlsm)`);
  }

  const excel = new ExcelJS.Workbook();
  const started = Date.now();
  await excel.xlsx.readFile(absPath);
  log.debug('excel.xlsx.readFile complete', { path: absPath, ms: Date.now() - started });

  const sheets: Record<string, HfCellValue[][]> = {};
  const sheetMeta: SheetMetadata[] = [];
  excel.worksheets.forEach((worksheet, index) => {
    const grid = extractSheetGrid(worksheet);
    sheets[worksheet.name] = grid;
    sheetMeta.push({
      name: worksheet.name,
      index,
      rowCount: worksheet.rowCount ?? 0,
      colCount: worksheet.columnCount ?? 0,
    });
  });

  const hfStarted = Date.now();
  const hf = HyperFormula.buildFromSheets(sheets, {
    licenseKey: 'gpl-v3',
  });
  log.debug('HyperFormula.buildFromSheets complete', { ms: Date.now() - hfStarted });

  const namedExpressions = extractNamedExpressions(excel, hf);

  const metadata: WorkbookMetadata = {
    id: opts.id,
    filePath: absPath,
    sizeBytes: stat.size,
    loadedAt: new Date().toISOString(),
    sheets: sheetMeta,
    namedExpressions,
  };

  return {
    id: opts.id,
    filePath: absPath,
    sizeBytes: stat.size,
    loadedAt: new Date(),
    hf,
    excel,
    metadata,
  };
}

function extractNamedExpressions(
  excel: ExcelJS.Workbook,
  hf: HyperFormula,
): WorkbookMetadata['namedExpressions'] {
  const out: WorkbookMetadata['namedExpressions'] = [];
  const dn = excel.definedNames as unknown as {
    matrixMap?: Record<string, unknown>;
    getRanges?: (name: string) => { name: string; ranges: string[] };
  };
  const map = dn?.matrixMap;
  if (!map || typeof map !== 'object') return out;
  for (const name of Object.keys(map)) {
    try {
      const info = dn.getRanges?.(name);
      const ranges = info?.ranges ?? [];
      if (ranges.length === 0) continue;
      const expression = ranges.length === 1 ? `=${ranges[0]}` : `={${ranges.join(',')}}`;
      try {
        hf.addNamedExpression(name, expression);
      } catch (e) {
        log.debug('skip named expression not accepted by HyperFormula', {
          name,
          expression,
          error: (e as Error).message,
        });
      }
      out.push({ name, expression, scope: null });
    } catch (e) {
      log.debug('failed to read named expression', { name, error: (e as Error).message });
    }
  }
  return out;
}

/** Guard against unlinking the file while a workbook is loaded. */
export function fileStillExists(filePath: string): boolean {
  try {
    return fs.statSync(filePath).isFile();
  } catch {
    return false;
  }
}
