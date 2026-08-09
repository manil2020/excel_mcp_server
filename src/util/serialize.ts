import { DetailedCellError } from 'hyperformula';

/** JSON-safe representation of a HyperFormula cell value. */
export type SerializedCellValue =
  | string
  | number
  | boolean
  | null
  | { error: string; message?: string; type?: string };

export function serializeCellValue(v: unknown): SerializedCellValue {
  if (v === null || v === undefined) return null;
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') {
    if (typeof v === 'number' && !Number.isFinite(v)) return null;
    return v;
  }
  if (v instanceof DetailedCellError) {
    return { error: v.value, message: v.message, type: v.type };
  }
  if (typeof v === 'object' && v !== null) {
    const anyV = v as { value?: unknown; type?: unknown; message?: unknown };
    if (typeof anyV.value === 'string' && typeof anyV.type === 'string') {
      return { error: anyV.value, message: typeof anyV.message === 'string' ? anyV.message : undefined, type: anyV.type };
    }
  }
  return String(v);
}
