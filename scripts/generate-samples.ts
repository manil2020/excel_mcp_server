import ExcelJS from 'exceljs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outDir = path.join(__dirname, '..', 'samples');

async function writeFinanceQuarterly(): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'excel-mcp-server samples';
  wb.created = new Date('2026-01-01T00:00:00Z');

  const regions = ['North', 'South', 'East', 'West'];
  const productLines = ['Alpha', 'Bravo', 'Charlie', 'Delta'];
  const quarters: [string, number][] = [
    ['Q1', 1_000_000],
    ['Q2', 1_200_000],
    ['Q3', 950_000],
    ['Q4', 1_450_000],
  ];

  for (const [q, base] of quarters) {
    const sheet = wb.addWorksheet(q);
    sheet.columns = [
      { header: 'Region', key: 'region', width: 12 },
      { header: 'Product', key: 'product', width: 12 },
      { header: 'Units', key: 'units', width: 10 },
      { header: 'UnitPrice', key: 'price', width: 12, style: { numFmt: '"$"#,##0.00' } },
      { header: 'Revenue', key: 'revenue', width: 14, style: { numFmt: '"$"#,##0.00' } },
      { header: 'Cost', key: 'cost', width: 14, style: { numFmt: '"$"#,##0.00' } },
      { header: 'Profit', key: 'profit', width: 14, style: { numFmt: '"$"#,##0.00' } },
    ];
    let row = 2;
    for (const region of regions) {
      for (const product of productLines) {
        const units = Math.round(base / (regions.length * productLines.length) / 100) + row * 3;
        const price = 99.99 + (row % 5) * 25;
        sheet.getCell(`A${row}`).value = region;
        sheet.getCell(`B${row}`).value = product;
        sheet.getCell(`C${row}`).value = units;
        sheet.getCell(`D${row}`).value = price;
        sheet.getCell(`E${row}`).value = { formula: `C${row}*D${row}` } as ExcelJS.CellFormulaValue;
        sheet.getCell(`F${row}`).value = { formula: `E${row}*0.6` } as ExcelJS.CellFormulaValue;
        sheet.getCell(`G${row}`).value = { formula: `E${row}-F${row}` } as ExcelJS.CellFormulaValue;
        row++;
      }
    }
    const totalRow = row;
    sheet.getCell(`A${totalRow}`).value = 'Total';
    sheet.getCell(`A${totalRow}`).font = { bold: true };
    sheet.getCell(`E${totalRow}`).value = { formula: `SUM(E2:E${totalRow - 1})` } as ExcelJS.CellFormulaValue;
    sheet.getCell(`F${totalRow}`).value = { formula: `SUM(F2:F${totalRow - 1})` } as ExcelJS.CellFormulaValue;
    sheet.getCell(`G${totalRow}`).value = { formula: `SUM(G2:G${totalRow - 1})` } as ExcelJS.CellFormulaValue;
  }

  const yearly = wb.addWorksheet('Yearly');
  yearly.getCell('A1').value = 'Quarter';
  yearly.getCell('B1').value = 'Revenue';
  yearly.getCell('C1').value = 'Cost';
  yearly.getCell('D1').value = 'Profit';
  yearly.getCell('E1').value = 'Margin %';
  yearly.getRow(1).font = { bold: true };
  quarters.forEach(([q], i) => {
    const r = i + 2;
    yearly.getCell(`A${r}`).value = q;
    const lastRow = regions.length * productLines.length + 1;
    yearly.getCell(`B${r}`).value = { formula: `${q}!E${lastRow}` } as ExcelJS.CellFormulaValue;
    yearly.getCell(`C${r}`).value = { formula: `${q}!F${lastRow}` } as ExcelJS.CellFormulaValue;
    yearly.getCell(`D${r}`).value = { formula: `${q}!G${lastRow}` } as ExcelJS.CellFormulaValue;
    yearly.getCell(`E${r}`).value = { formula: `D${r}/B${r}` } as ExcelJS.CellFormulaValue;
    yearly.getCell(`E${r}`).numFmt = '0.0%';
  });
  const grandRow = quarters.length + 2;
  yearly.getCell(`A${grandRow}`).value = 'Full Year';
  yearly.getCell(`A${grandRow}`).font = { bold: true };
  yearly.getCell(`B${grandRow}`).value = { formula: `SUM(B2:B${grandRow - 1})` } as ExcelJS.CellFormulaValue;
  yearly.getCell(`C${grandRow}`).value = { formula: `SUM(C2:C${grandRow - 1})` } as ExcelJS.CellFormulaValue;
  yearly.getCell(`D${grandRow}`).value = { formula: `SUM(D2:D${grandRow - 1})` } as ExcelJS.CellFormulaValue;
  yearly.getCell(`E${grandRow}`).value = { formula: `D${grandRow}/B${grandRow}` } as ExcelJS.CellFormulaValue;
  yearly.getCell(`E${grandRow}`).numFmt = '0.0%';

  wb.definedNames.add('Yearly!$B$6', 'FullYearRevenue');

  const out = path.join(outDir, 'finance-quarterly.xlsx');
  await wb.xlsx.writeFile(out);
  console.log(`  wrote ${out}`);
}

async function writeEmployeeHR(): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'excel-mcp-server samples';

  const employees = wb.addWorksheet('Employees');
  employees.columns = [
    { header: 'EmpID', key: 'id', width: 8 },
    { header: 'Name', key: 'name', width: 20 },
    { header: 'Department', key: 'dept', width: 14 },
    { header: 'HireDate', key: 'hire', width: 12, style: { numFmt: 'yyyy-mm-dd' } },
    { header: 'Salary', key: 'salary', width: 12, style: { numFmt: '"$"#,##0.00' } },
    { header: 'TenureYears', key: 'tenure', width: 12 },
    { header: 'Band', key: 'band', width: 8 },
  ];
  const names = [
    'Alice Chen', 'Bob Martinez', 'Carol Nguyen', 'Dave Patel', 'Eve Johnson',
    'Frank Kumar', 'Grace Wong', 'Henry Silva', 'Ivy Larsson', 'Jack Ohara',
    'Kira Bello', 'Liam Zhou', 'Mia Rossi', 'Noah Kim', 'Olivia Adekunle',
  ];
  const depts = ['Engineering', 'Sales', 'HR', 'Finance', 'Marketing'];
  names.forEach((name, i) => {
    const r = i + 2;
    const hireYear = 2015 + (i % 10);
    const hireDate = new Date(hireYear, i % 12, ((i * 7) % 27) + 1);
    const dept = depts[i % depts.length];
    const salary = 55_000 + (i * 3_500);
    employees.getCell(`A${r}`).value = 1000 + i;
    employees.getCell(`B${r}`).value = name;
    employees.getCell(`C${r}`).value = dept;
    employees.getCell(`D${r}`).value = hireDate;
    employees.getCell(`E${r}`).value = salary;
    employees.getCell(`F${r}`).value = { formula: `DATEDIF(D${r},TODAY(),"Y")` } as ExcelJS.CellFormulaValue;
    employees.getCell(`G${r}`).value = {
      formula: `IF(E${r}<70000,"B1",IF(E${r}<90000,"B2",IF(E${r}<110000,"B3","B4")))`,
    } as ExcelJS.CellFormulaValue;
  });

  const bands = wb.addWorksheet('SalaryBands');
  bands.columns = [
    { header: 'Band', key: 'band', width: 8 },
    { header: 'MinSalary', key: 'min', width: 12 },
    { header: 'MaxSalary', key: 'max', width: 12 },
    { header: 'Description', key: 'desc', width: 30 },
  ];
  const bandData: [string, number, number, string][] = [
    ['B1', 40_000, 69_999, 'Junior individual contributor'],
    ['B2', 70_000, 89_999, 'Mid-level individual contributor'],
    ['B3', 90_000, 109_999, 'Senior individual contributor'],
    ['B4', 110_000, 200_000, 'Staff / Manager+'],
  ];
  bandData.forEach((row, i) => {
    const r = i + 2;
    bands.getCell(`A${r}`).value = row[0];
    bands.getCell(`B${r}`).value = row[1];
    bands.getCell(`C${r}`).value = row[2];
    bands.getCell(`D${r}`).value = row[3];
  });

  const summary = wb.addWorksheet('Summary');
  summary.getCell('A1').value = 'Department';
  summary.getCell('B1').value = 'HeadCount';
  summary.getCell('C1').value = 'TotalPayroll';
  summary.getCell('D1').value = 'AvgSalary';
  summary.getRow(1).font = { bold: true };
  depts.forEach((d, i) => {
    const r = i + 2;
    summary.getCell(`A${r}`).value = d;
    summary.getCell(`B${r}`).value = { formula: `COUNTIF(Employees!C:C,A${r})` } as ExcelJS.CellFormulaValue;
    summary.getCell(`C${r}`).value = { formula: `SUMIF(Employees!C:C,A${r},Employees!E:E)` } as ExcelJS.CellFormulaValue;
    summary.getCell(`D${r}`).value = { formula: `IFERROR(C${r}/B${r},0)` } as ExcelJS.CellFormulaValue;
    summary.getCell(`C${r}`).numFmt = '"$"#,##0.00';
    summary.getCell(`D${r}`).numFmt = '"$"#,##0.00';
  });

  const out = path.join(outDir, 'employee-data.xlsx');
  await wb.xlsx.writeFile(out);
  console.log(`  wrote ${out}`);
}

async function writeMessyReport(): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'excel-mcp-server samples';

  const sheet = wb.addWorksheet('Monthly Report');
  sheet.mergeCells('A1:E1');
  sheet.getCell('A1').value = 'ACME Corp — Monthly Sales Report';
  sheet.getCell('A1').font = { bold: true, size: 14 };
  sheet.getCell('A1').alignment = { horizontal: 'center' };

  sheet.getCell('A3').value = 'Prepared by:';
  sheet.getCell('B3').value = 'Finance Team';
  sheet.getCell('D3').value = 'Period:';
  sheet.getCell('E3').value = 'Feb 2026';

  // Real headers on row 5 (typical "messy Excel" pattern)
  const headers = ['SKU', 'Product', 'Category', 'Units Sold', 'Unit Price', 'Line Total'];
  headers.forEach((h, i) => {
    const c = String.fromCharCode(65 + i);
    sheet.getCell(`${c}5`).value = h;
    sheet.getCell(`${c}5`).font = { bold: true };
    sheet.getCell(`${c}5`).border = { bottom: { style: 'thin' } };
  });

  const rows: [string, string, string, number, number][] = [
    ['SKU-001', 'Widget Deluxe', 'Widgets', 45, 29.99],
    ['SKU-002', 'Gizmo Pro', 'Gizmos', 12, 149.5],
    ['SKU-003', 'Sprocket XL', 'Sprockets', 88, 8.75],
    ['SKU-004', 'Thingamajig', 'Widgets', 0, 55.0],
    ['SKU-005', 'Doohickey', 'Gizmos', 33, 22.5],
  ];
  rows.forEach((row, i) => {
    const r = i + 6;
    sheet.getCell(`A${r}`).value = row[0];
    sheet.getCell(`B${r}`).value = row[1];
    sheet.getCell(`C${r}`).value = row[2];
    sheet.getCell(`D${r}`).value = row[3];
    sheet.getCell(`E${r}`).value = row[4];
    sheet.getCell(`F${r}`).value = { formula: `D${r}*E${r}` } as ExcelJS.CellFormulaValue;
    sheet.getCell(`E${r}`).numFmt = '"$"#,##0.00';
    sheet.getCell(`F${r}`).numFmt = '"$"#,##0.00';
  });

  const totalRow = rows.length + 7;
  sheet.getCell(`E${totalRow}`).value = 'Grand Total';
  sheet.getCell(`E${totalRow}`).font = { bold: true };
  sheet.getCell(`F${totalRow}`).value = { formula: `SUM(F6:F${rows.length + 5})` } as ExcelJS.CellFormulaValue;
  sheet.getCell(`F${totalRow}`).font = { bold: true };
  sheet.getCell(`F${totalRow}`).numFmt = '"$"#,##0.00';

  // Notes sheet — free-form text intentional
  const notes = wb.addWorksheet('Notes');
  notes.getCell('A1').value = 'Notes';
  notes.getCell('A1').font = { bold: true, size: 14 };
  notes.getCell('A3').value = 'SKU-004 had a manufacturing recall; 0 units shipped this month.';
  notes.getCell('A4').value = 'Widget Deluxe is the top seller for the second month running.';
  notes.getCell('A5').value = 'Follow-up call scheduled with the Gizmos supplier on 2026-02-28.';

  const out = path.join(outDir, 'messy-report.xlsx');
  await wb.xlsx.writeFile(out);
  console.log(`  wrote ${out}`);
}

async function writeLargeDataset(): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'excel-mcp-server samples';
  const sheet = wb.addWorksheet('Transactions');
  sheet.columns = [
    { header: 'TxID', key: 'id', width: 10 },
    { header: 'Date', key: 'date', width: 12, style: { numFmt: 'yyyy-mm-dd' } },
    { header: 'Customer', key: 'customer', width: 18 },
    { header: 'Country', key: 'country', width: 12 },
    { header: 'Product', key: 'product', width: 18 },
    { header: 'Quantity', key: 'qty', width: 10 },
    { header: 'UnitPrice', key: 'price', width: 12 },
    { header: 'Total', key: 'total', width: 14 },
  ];
  const customers = ['Acme', 'Globex', 'Initech', 'Umbrella', 'Soylent', 'Cyberdyne', 'Tyrell', 'Weyland'];
  const countries = ['US', 'UK', 'DE', 'FR', 'JP', 'IN', 'BR', 'AU'];
  const products = ['Widget-A', 'Widget-B', 'Gizmo-C', 'Sprocket-D', 'Doohickey-E'];
  const startDate = new Date(2025, 0, 1).getTime();
  const N = 2_000;
  for (let i = 0; i < N; i++) {
    const r = i + 2;
    const date = new Date(startDate + i * 3_600_000 * 8);
    const qty = 1 + ((i * 13) % 20);
    const price = 9.99 + ((i * 7) % 90);
    sheet.getCell(`A${r}`).value = 10_000 + i;
    sheet.getCell(`B${r}`).value = date;
    sheet.getCell(`C${r}`).value = customers[i % customers.length];
    sheet.getCell(`D${r}`).value = countries[i % countries.length];
    sheet.getCell(`E${r}`).value = products[i % products.length];
    sheet.getCell(`F${r}`).value = qty;
    sheet.getCell(`G${r}`).value = price;
    sheet.getCell(`H${r}`).value = { formula: `F${r}*G${r}` } as ExcelJS.CellFormulaValue;
  }

  const out = path.join(outDir, 'transactions-2k.xlsx');
  await wb.xlsx.writeFile(out);
  console.log(`  wrote ${out}`);
}

async function main(): Promise<void> {
  console.log('Generating sample workbooks in', outDir);
  await writeFinanceQuarterly();
  await writeEmployeeHR();
  await writeMessyReport();
  await writeLargeDataset();
  console.log('done');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
