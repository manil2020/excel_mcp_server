import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { createServer } from '../src/server.js';
import { WorkbookManager } from '../src/workbook/manager.js';
import { createSampleXlsx, type SampleFixture } from './fixtures/create-sample.js';

let fixture: SampleFixture;
let client: Client;
let manager: WorkbookManager;

beforeAll(async () => {
  fixture = await createSampleXlsx();
  manager = new WorkbookManager();
  const { server } = createServer({ manager, name: 'excel-mcp-test', version: '0.0.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  client = new Client({ name: 'excel-mcp-test-client', version: '0.0.0' });
  await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);
});

afterAll(async () => {
  await client.close();
  manager.closeAll();
  await fixture.cleanup();
});

interface StructuredResult {
  structuredContent?: Record<string, unknown>;
  isError?: boolean;
  content: { type: string; text?: string }[];
}

async function call(name: string, args: Record<string, unknown> = {}): Promise<StructuredResult> {
  const raw = await client.callTool({ name, arguments: args });
  return raw as unknown as StructuredResult;
}

describe('MCP tools', () => {
  it('lists all registered tools', async () => {
    const list = await client.listTools();
    const names = list.tools.map((t) => t.name).sort();
    expect(names).toEqual(
      [
        'close_workbook',
        'evaluate_formula',
        'find_in_workbook',
        'get_cell',
        'get_sheet_summary',
        'get_workbook_stats',
        'list_sheets',
        'list_workbooks',
        'open_workbook',
        'read_range',
      ].sort(),
    );
  });

  it('open_workbook returns a workbookId and sheet list', async () => {
    const res = await call('open_workbook', { path: fixture.filePath });
    expect(res.isError).toBeFalsy();
    const s = res.structuredContent as { workbookId: string; sheets: { name: string }[] };
    expect(s.workbookId).toMatch(/^wb_/);
    expect(s.sheets.map((sh) => sh.name)).toEqual(['Data', 'Summary']);
  });

  it('supports the open → read_range → evaluate_formula → close flow', async () => {
    const opened = await call('open_workbook', { path: fixture.filePath });
    const id = (opened.structuredContent as { workbookId: string }).workbookId;

    const range = await call('read_range', {
      workbookId: id,
      sheet: 'Data',
      range: 'A1:D5',
      includeFormulas: true,
    });
    expect(range.isError).toBeFalsy();
    const rangePayload = range.structuredContent as {
      values: (string | number | null)[][];
      formulas: (string | null)[][];
    };
    expect(rangePayload.values[0]).toEqual(['id', 'name', 'amount', 'double']);
    expect(rangePayload.values[1]).toEqual([1, 'Alice', 10, 20]);
    expect(rangePayload.formulas[1]?.[3]).toContain('C2*2');

    const cell = await call('get_cell', { workbookId: id, sheet: 'Summary', cell: 'B1' });
    const cellPayload = cell.structuredContent as { value: number; formula: string };
    expect(cellPayload.value).toBe(84);
    expect(cellPayload.formula).toBe('=SUM(Data!C2:C5)');

    const evalRes = await call('evaluate_formula', {
      workbookId: id,
      formula: '=SUM(Data!D2:D5)/2',
      sheet: 'Summary',
    });
    expect((evalRes.structuredContent as { result: number }).result).toBe(84);

    const stats = await call('get_workbook_stats', { workbookId: id });
    const statsPayload = stats.structuredContent as { totals: { formulaCells: number } };
    expect(statsPayload.totals.formulaCells).toBeGreaterThanOrEqual(7);

    const closed = await call('close_workbook', { workbookId: id });
    expect((closed.structuredContent as { closed: boolean }).closed).toBe(true);
  });

  it('find_in_workbook locates values across sheets', async () => {
    const opened = await call('open_workbook', { path: fixture.filePath });
    const id = (opened.structuredContent as { workbookId: string }).workbookId;

    const res = await call('find_in_workbook', { workbookId: id, query: 'Carol' });
    const payload = res.structuredContent as {
      hits: { sheet: string; cell: string; matchedIn: string }[];
      hitCount: number;
    };
    expect(payload.hitCount).toBeGreaterThanOrEqual(1);
    expect(payload.hits[0]?.sheet).toBe('Data');

    await call('close_workbook', { workbookId: id });
  });

  it('returns a structured error for unknown workbookId', async () => {
    const res = await call('list_sheets', { workbookId: 'wb_deadbeef' });
    expect(res.isError).toBe(true);
    const payload = res.structuredContent as { error: string; message: string };
    expect(payload.error).toBe('WORKBOOK_NOT_FOUND');
    expect(payload.message).toMatch(/wb_deadbeef/);
  });
});
