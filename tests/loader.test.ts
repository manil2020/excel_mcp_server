import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { WorkbookManager } from '../src/workbook/manager.js';
import { createSampleXlsx, type SampleFixture } from './fixtures/create-sample.js';

let fixture: SampleFixture;
let manager: WorkbookManager;

beforeAll(async () => {
  fixture = await createSampleXlsx();
  manager = new WorkbookManager();
});

afterAll(async () => {
  manager.closeAll();
  await fixture.cleanup();
});

describe('loader + manager', () => {
  it('opens a workbook and enumerates sheets', async () => {
    const wb = await manager.open(fixture.filePath);
    expect(wb.id).toMatch(/^wb_/);
    expect(wb.metadata.sheets.map((s) => s.name)).toEqual(['Data', 'Summary']);
    const data = wb.metadata.sheets.find((s) => s.name === 'Data');
    expect(data?.rowCount).toBeGreaterThanOrEqual(5);
    expect(data?.colCount).toBeGreaterThanOrEqual(4);
    manager.close(wb.id);
  });

  it('evaluates original workbook formulas through HyperFormula', async () => {
    const wb = await manager.open(fixture.filePath);
    const summaryId = wb.hf.getSheetId('Summary')!;
    // Total = 10 + 25 + 42 + 7 = 84
    expect(wb.hf.getCellValue({ sheet: summaryId, row: 0, col: 1 })).toBe(84);
    // Max = 42
    expect(wb.hf.getCellValue({ sheet: summaryId, row: 1, col: 1 })).toBe(42);
    // Doubled total = 168
    expect(wb.hf.getCellValue({ sheet: summaryId, row: 2, col: 1 })).toBe(168);
    manager.close(wb.id);
  });

  it('evaluates ad-hoc formulas via calculateFormula', async () => {
    const wb = await manager.open(fixture.filePath);
    const dataId = wb.hf.getSheetId('Data')!;
    expect(wb.hf.calculateFormula('=AVERAGE(C2:C5)', dataId)).toBe(21);
    expect(wb.hf.calculateFormula('=COUNTA(B2:B5)', dataId)).toBe(4);
    manager.close(wb.id);
  });

  it('enforces the maxWorkbooks limit', async () => {
    const bounded = new WorkbookManager({ maxWorkbooks: 1 });
    const first = await bounded.open(fixture.filePath);
    await expect(bounded.open(fixture.filePath)).rejects.toThrow(/Too many open workbooks/);
    bounded.close(first.id);
  });

  it('rejects unsupported file extensions', async () => {
    await expect(manager.open('/nonexistent/file.csv')).rejects.toThrow();
  });
});
