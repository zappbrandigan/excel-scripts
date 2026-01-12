/// <reference path="../office-scripts-docs.d.ts" />

/**
 * ---------------------------------------------------------------------------
 * Company: Universal Music Publishing Group
 * Script Name: Split Column A to Sheets
 * Author: Brandon Johnson
 * Date: 2026-01-11
 * Version: 1.0.0
 * ---------------------------------------------------------------------------
 * Description:
 * Splits values from column A into multiple worksheets, capped by a maximum
 * rows-per-sheet value. Optionally targets a specific worksheet and can skip
 * the header row.
 *
 * Features:
 * - Detects row count from column A or the overall used range.
 * - Splits data into new sheets named Split_01, Split_02, etc.
 * - Copies only column A values to keep output minimal.
 * - Auto-fits the output column for readability.
 *
 * Parameters:
 * - maxRowsPerSheet: number
 *   Maximum number of data rows per output worksheet (default: 30000).
 * - sheetName?: string
 *   Optional worksheet name to target instead of the active sheet.
 * - skipHeader: boolean
 *   Whether to ignore the first row as a header (default: true).
 *
 * Notes:
 * - Only column A is copied to the new sheets.
 * - Existing worksheets are not modified; new ones are added.
 */
function main(
  workbook: ExcelScript.Workbook,
  maxRowsPerSheet: number = 30000,
  sheetName?: string,
  skipHeader: boolean = true
): void {
  const ws = sheetName
    ? workbook.getWorksheet(sheetName)
    : workbook.getActiveWorksheet();
  if (!ws) {
    console.log('No worksheet found.');
    return;
  }

  // 1) Prefer the used range of column A; fall back to whole-sheet used range
  const usedA = ws.getRange('A:A').getUsedRange(); // may be null
  const usedAll = ws.getUsedRange(); // may be null
  const rowCountA = usedA ? usedA.getRowCount() : 0;
  const rowCountAll = usedAll ? usedAll.getRowCount() : 0;

  // Heuristic: if column A has something, trust it; otherwise trust whole-sheet
  const totalRows = Math.max(rowCountA, rowCountAll);

  if (!totalRows || totalRows <= (skipHeader ? 1 : 0)) {
    console.log('No values detected in column A (or only a header).');
    return;
  }

  // 2) Read only column A, from row 1..totalRows
  //    (Indexes are 0-based; column A is index 0)
  const firstDataRowIndex = skipHeader ? 1 : 0;
  const dataRowCount = totalRows - firstDataRowIndex;

  const colARange = ws.getRangeByIndexes(firstDataRowIndex, 0, dataRowCount, 1);
  const colAValues = colARange.getValues() as (
    | string
    | number
    | boolean
    | null
  )[][];
  if (!colAValues || colAValues.length === 0) {
    console.log('Column A has no readable values in the detected range.');
    return;
  }

  // 3) Coerce null/undefined to empty strings for safe setValues()
  const cleaned: (string | number | boolean)[][] = colAValues.map((r) => [
    coerce(r?.[0]),
  ]);

  // 4) Split into sheets of ≤ maxRowsPerSheet
  const total = cleaned.length;
  const chunks = Math.ceil(total / maxRowsPerSheet);

  for (let i = 0; i < chunks; i++) {
    const start = i * maxRowsPerSheet;
    const end = Math.min(start + maxRowsPerSheet, total);
    const slice = cleaned.slice(start, end);

    const newWs = workbook.addWorksheet(
      `Split_${String(i + 1).padStart(2, '0')}`
    );
    const outRange = newWs.getRangeByIndexes(0, 0, slice.length, 1);
    outRange.setValues(slice);
    outRange.getFormat().autofitColumns();
  }

  function coerce(
    v: string | number | boolean | null | undefined
  ): string | number | boolean {
    return v === null || v === undefined ? '' : v;
  }
}
