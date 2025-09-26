/// <reference path="../office-scripts-docs.d.ts" />

function main(
  workbook: ExcelScript.Workbook,
  maxRowsPerSheet: number = 30000
): void {
  const ws = workbook.getActiveWorksheet();

  const used = ws.getUsedRange();
  if (!used) {
    console.log('No used range on the active sheet.');
    return;
  }

  // Values may contain nulls for empty cells; type accordingly
  const values = used.getValues() as (string | number | boolean | null)[][];
  if (!values || values.length === 0) {
    console.log('No values in used range.');
    return;
  }

  // --- First column only, skip the first row (no headers)
  // Coerce nulls to empty string so setValues() is happy
  const firstCol: (string | number | boolean)[][] = values
    .slice(1)
    .map((row) => [coerce(row?.[0])]);

  if (firstCol.length === 0) {
    console.log('No data rows after skipping header.');
    return;
  }

  const total = firstCol.length;
  const chunks = Math.ceil(total / maxRowsPerSheet);

  for (let i = 0; i < chunks; i++) {
    const start = i * maxRowsPerSheet;
    const end = Math.min(start + maxRowsPerSheet, total);
    const slice = firstCol.slice(start, end);

    const newWsName = `Split_${String(i + 1).padStart(2, '0')}`;
    const newWs = workbook.addWorksheet(newWsName);

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
