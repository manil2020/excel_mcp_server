import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ExcelMcpError } from '../util/errors.js';

export function ok(payload: Record<string, unknown>): CallToolResult {
  return {
    content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }],
    structuredContent: payload,
  };
}

export function fail(err: unknown): CallToolResult {
  if (err instanceof ExcelMcpError) {
    return {
      isError: true,
      content: [
        {
          type: 'text',
          text: JSON.stringify({ error: err.code, message: err.message, details: err.details }, null, 2),
        },
      ],
      structuredContent: { error: err.code, message: err.message, details: err.details ?? {} },
    };
  }
  const message = err instanceof Error ? err.message : String(err);
  return {
    isError: true,
    content: [{ type: 'text', text: JSON.stringify({ error: 'INTERNAL', message }, null, 2) }],
    structuredContent: { error: 'INTERNAL', message },
  };
}
