import * as path from 'node:path';
import * as fs from 'node:fs';
import { FileAccessError } from './errors.js';

function parseAllowedRoots(): string[] | null {
  const raw = process.env.EXCEL_MCP_ALLOWED_ROOTS;
  if (!raw) return null;
  return raw
    .split(':')
    .map((p) => p.trim())
    .filter(Boolean)
    .map((p) => path.resolve(p));
}

/**
 * Resolve a user-provided path to an absolute path and enforce the optional
 * allow-list configured via EXCEL_MCP_ALLOWED_ROOTS. When no allow-list is set,
 * any readable local path is accepted.
 */
export function resolveSafePath(inputPath: string): string {
  if (!inputPath || typeof inputPath !== 'string') {
    throw new FileAccessError(String(inputPath), 'path must be a non-empty string');
  }
  const abs = path.resolve(inputPath);
  const allowed = parseAllowedRoots();
  if (allowed && allowed.length > 0) {
    const within = allowed.some((root) => {
      const rel = path.relative(root, abs);
      return rel === '' || (!rel.startsWith('..') && !path.isAbsolute(rel));
    });
    if (!within) {
      throw new FileAccessError(
        abs,
        `path is outside the allowed roots (${allowed.join(', ')})`,
      );
    }
  }
  return abs;
}

export function assertReadableFile(absPath: string): fs.Stats {
  let stat: fs.Stats;
  try {
    stat = fs.statSync(absPath);
  } catch (e) {
    throw new FileAccessError(absPath, (e as Error).message);
  }
  if (!stat.isFile()) {
    throw new FileAccessError(absPath, 'not a regular file');
  }
  try {
    fs.accessSync(absPath, fs.constants.R_OK);
  } catch (e) {
    throw new FileAccessError(absPath, `not readable (${(e as Error).message})`);
  }
  return stat;
}
