import ExcelJS from 'exceljs';
import * as fs from 'node:fs/promises';
import * as os from 'node:os';
import * as path from 'node:path';

export interface SampleFixture {
  filePath: string;
  cleanup: () => Promise<void>;
}

/**
 * Write a small deterministic .xlsx to a temp dir. Returns the file path plus
 * a cleanup callback the caller should invoke in afterAll.
 */
export async function createSampleXlsx(): Promise<SampleFixture> {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'excel-mcp-test-'));
  const filePath = path.join(dir, 'sample.xlsx');

  const wb = new ExcelJS.Workbook();
  wb.creator = 'excel-mcp-server tests';
  wb.created = new Date('2026-01-01T00:00:00Z');

  const data = wb.addWorksheet('Data');
  data.columns = [
    { header: 'id', key: 'id', width: 6 },
    { header: 'name', key: 'name', width: 12 },
    { header: 'amount', key: 'amount', width: 10 },
    { header: 'double', key: 'double', width: 10 },
  ];
  data.addRow({ id: 1, name: 'Alice', amount: 10 });
  data.addRow({ id: 2, name: 'Bob', amount: 25 });
  data.addRow({ id: 3, name: 'Carol', amount: 42 });
  data.addRow({ id: 4, name: 'Dave', amount: 7 });
  for (let r = 2; r <= 5; r++) {
    data.getCell(`D${r}`).value = { formula: `C${r}*2`, date1904: false } as unknown as ExcelJS.CellFormulaValue;
  }

  const summary = wb.addWorksheet('Summary');
  summary.getCell('A1').value = 'Total';
  summary.getCell('B1').value = { formula: 'SUM(Data!C2:C5)', date1904: false } as unknown as ExcelJS.CellFormulaValue;
  summary.getCell('A2').value = 'Max';
  summary.getCell('B2').value = { formula: 'MAX(Data!C2:C5)', date1904: false } as unknown as ExcelJS.CellFormulaValue;
  summary.getCell('A3').value = 'Doubled Total';
  summary.getCell('B3').value = { formula: 'SUM(Data!D2:D5)', date1904: false } as unknown as ExcelJS.CellFormulaValue;

  await wb.xlsx.writeFile(filePath);

  return {
    filePath,
    cleanup: async () => {
      try {
        await fs.rm(dir, { recursive: true, force: true });
      } catch {
        /* ignore */
      }
    },
  };
}
