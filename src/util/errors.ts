export class ExcelMcpError extends Error {
  constructor(
    message: string,
    public readonly code: string,
    public readonly details?: Record<string, unknown>,
  ) {
    super(message);
    this.name = 'ExcelMcpError';
  }
}

export class WorkbookNotFoundError extends ExcelMcpError {
  constructor(id: string) {
    super(`No open workbook with id "${id}". Call open_workbook first.`, 'WORKBOOK_NOT_FOUND', { id });
  }
}

export class SheetNotFoundError extends ExcelMcpError {
  constructor(sheet: string, available: string[]) {
    super(
      `Sheet "${sheet}" not found. Available sheets: ${available.join(', ')}`,
      'SHEET_NOT_FOUND',
      { sheet, available },
    );
  }
}

export class InvalidRangeError extends ExcelMcpError {
  constructor(range: string, reason: string) {
    super(`Invalid range "${range}": ${reason}`, 'INVALID_RANGE', { range, reason });
  }
}

export class FileAccessError extends ExcelMcpError {
  constructor(path: string, reason: string) {
    super(`Cannot access file "${path}": ${reason}`, 'FILE_ACCESS', { path, reason });
  }
}

export class FileTooLargeError extends ExcelMcpError {
  constructor(path: string, sizeBytes: number, maxBytes: number) {
    super(
      `File "${path}" is ${(sizeBytes / 1024 / 1024).toFixed(1)} MB which exceeds the configured limit of ${(maxBytes / 1024 / 1024).toFixed(0)} MB. Set EXCEL_MCP_MAX_FILE_MB to raise the limit.`,
      'FILE_TOO_LARGE',
      { path, sizeBytes, maxBytes },
    );
  }
}

export class UnsupportedFileError extends ExcelMcpError {
  constructor(path: string, reason: string) {
    super(`Unsupported file "${path}": ${reason}`, 'UNSUPPORTED_FILE', { path, reason });
  }
}

export class TooManyWorkbooksError extends ExcelMcpError {
  constructor(current: number, max: number) {
    super(
      `Too many open workbooks (${current}/${max}). Close one via close_workbook first, or raise EXCEL_MCP_MAX_WORKBOOKS.`,
      'TOO_MANY_WORKBOOKS',
      { current, max },
    );
  }
}
