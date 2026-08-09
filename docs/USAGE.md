# Usage guide

Every tool takes JSON arguments and returns JSON-serialisable output. All examples below show the *arguments* you pass and the *shape* of the result. Field names in the result are stable.

## `open_workbook`

```json
{ "path": "/data/orders.xlsx" }
```

Returns:

```json
{
  "workbookId": "wb_ab12cd34",
  "filePath": "/data/orders.xlsx",
  "sizeBytes": 45231,
  "sheets": [
    { "name": "Data",    "index": 0, "rowCount": 100, "colCount": 6 },
    { "name": "Summary", "index": 1, "rowCount": 3,   "colCount": 2 }
  ],
  "namedExpressions": [ { "name": "TaxRate", "expression": "=Config!$B$2", "scope": null } ],
  "namedExpressionsTruncated": false
}
```

## `list_workbooks`

No arguments. Returns every currently open workbook.

## `list_sheets`

```json
{ "workbookId": "wb_ab12cd34" }
```

## `get_sheet_summary`

```json
{ "workbookId": "wb_ab12cd34", "sheet": "Data", "sampleRows": 20 }
```

Returns dimensions, detected header row (if the first row looks like column labels), inferred column types, and a preview of the first `sampleRows` rows.

## `read_range`

```json
{
  "workbookId": "wb_ab12cd34",
  "sheet": "Data",
  "range": "A1:D50",
  "includeFormulas": true,
  "maxCells": 5000
}
```

The result contains `values` (row-major 2D array of evaluated values) and, when `includeFormulas` is `true`, a parallel `formulas` array with the original formula strings (`null` for non-formula cells).

## `get_cell`

```json
{ "workbookId": "wb_ab12cd34", "sheet": "Summary", "cell": "B1" }
```

Returns:

```json
{
  "value": 84,
  "formula": "SUM(Data!C2:C5)",
  "cellType": "FORMULA",
  "valueType": "NUMBER",
  "numberFormat": "General",
  "excelRaw": { "formula": "SUM(Data!C2:C5)", "result": 84 }
}
```

## `evaluate_formula`

Evaluate an arbitrary formula in the context of a loaded workbook.

```json
{
  "workbookId": "wb_ab12cd34",
  "formula": "=SUMIFS(Data!C:C, Data!D:D, \"EMEA\")",
  "sheet": "Data"
}
```

The `sheet` argument only affects how *relative* references resolve. Cross-sheet references (`Data!C:C` above) work regardless.

Array formulas return `resultType: "range"` and a 2D `result` array; scalars return `resultType: "scalar"`.

## `find_in_workbook`

```json
{
  "workbookId": "wb_ab12cd34",
  "query": "EMEA",
  "regex": false,
  "caseSensitive": false,
  "searchFormulas": true,
  "sheets": ["Orders", "Refunds"],
  "maxHits": 200
}
```

Returns each hit with sheet name, A1 cell address, evaluated value, formula (if any), and whether the match was in `value` or `formula`.

## `get_workbook_stats`

```json
{ "workbookId": "wb_ab12cd34" }
```

Returns per-sheet and total counts of non-empty cells, formulas, and error cells. Useful for the LLM to size its exploration strategy.

## `close_workbook`

```json
{ "workbookId": "wb_ab12cd34" }
```

Idempotent — returns `{ "closed": false }` if the id was already released.

---

## Error shape

Errors come back with `isError: true` and a structured payload:

```json
{
  "error": "SHEET_NOT_FOUND",
  "message": "Sheet \"Foo\" not found. Available sheets: Data, Summary",
  "details": { "sheet": "Foo", "available": ["Data", "Summary"] }
}
```

Known error codes: `WORKBOOK_NOT_FOUND`, `SHEET_NOT_FOUND`, `INVALID_RANGE`, `FILE_ACCESS`, `FILE_TOO_LARGE`, `UNSUPPORTED_FILE`, `TOO_MANY_WORKBOOKS`, `INTERNAL`.
