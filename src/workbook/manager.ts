import { randomBytes } from 'node:crypto';
import { loadWorkbook } from './loader.js';
import type { LoadedWorkbook, WorkbookMetadata } from './types.js';
import { TooManyWorkbooksError, WorkbookNotFoundError } from '../util/errors.js';
import { log } from '../util/logger.js';

const DEFAULT_MAX_WORKBOOKS = 8;
const DEFAULT_MAX_FILE_MB = 100;

function envInt(key: string, fallback: number): number {
  const raw = process.env[key];
  if (!raw) return fallback;
  const n = Number.parseInt(raw, 10);
  return Number.isFinite(n) && n > 0 ? n : fallback;
}

export class WorkbookManager {
  private readonly books = new Map<string, LoadedWorkbook>();
  private readonly maxWorkbooks: number;
  private readonly maxFileBytes: number;

  constructor(opts?: { maxWorkbooks?: number; maxFileBytes?: number }) {
    this.maxWorkbooks = opts?.maxWorkbooks ?? envInt('EXCEL_MCP_MAX_WORKBOOKS', DEFAULT_MAX_WORKBOOKS);
    this.maxFileBytes =
      opts?.maxFileBytes ?? envInt('EXCEL_MCP_MAX_FILE_MB', DEFAULT_MAX_FILE_MB) * 1024 * 1024;
  }

  async open(filePath: string): Promise<LoadedWorkbook> {
    if (this.books.size >= this.maxWorkbooks) {
      throw new TooManyWorkbooksError(this.books.size, this.maxWorkbooks);
    }
    const id = `wb_${randomBytes(4).toString('hex')}`;
    const wb = await loadWorkbook(filePath, { id, maxFileBytes: this.maxFileBytes });
    this.books.set(id, wb);
    log.info('workbook opened', {
      id,
      path: wb.filePath,
      sizeBytes: wb.sizeBytes,
      sheets: wb.metadata.sheets.length,
    });
    return wb;
  }

  get(id: string): LoadedWorkbook {
    const wb = this.books.get(id);
    if (!wb) throw new WorkbookNotFoundError(id);
    return wb;
  }

  close(id: string): boolean {
    const wb = this.books.get(id);
    if (!wb) return false;
    try {
      wb.hf.destroy();
    } catch (e) {
      log.warn('HyperFormula.destroy threw', { id, error: (e as Error).message });
    }
    this.books.delete(id);
    log.info('workbook closed', { id });
    return true;
  }

  list(): WorkbookMetadata[] {
    return [...this.books.values()].map((wb) => wb.metadata);
  }

  closeAll(): void {
    for (const id of [...this.books.keys()]) {
      this.close(id);
    }
  }

  get size(): number {
    return this.books.size;
  }

  get limits(): { maxWorkbooks: number; maxFileBytes: number } {
    return { maxWorkbooks: this.maxWorkbooks, maxFileBytes: this.maxFileBytes };
  }
}

export const defaultManager = new WorkbookManager();
