/// <reference path="../office-scripts-docs.d.ts" />

function main(
  workbook: ExcelScript.Workbook,
  maxRowsPerSheet: number = 30000
): void {
  const ws = workbook.getActiveWorksheet();
  const used = ws.getUsedRange(true);
  if (!used) return;

  // Get all values and pull out only first column (skip header entirely)
  const values = used.getValues();
  if (values.length === 0) return;

  const firstCol: any[][] = values.map((r) => [r[0]]); // 2D array with only first col
  const total = firstCol.length;
  const chunks = Math.ceil(total / maxRowsPerSheet);

  for (let i = 0; i < chunks; i++) {
    const start = i * maxRowsPerSheet;
    const end = Math.min(start + maxRowsPerSheet, total);
    const slice = firstCol.slice(start, end);

    // Create new sheet and drop values
    const newWs = workbook.addWorksheet(
      `Split_${String(i + 1).padStart(2, '0')}`
    );
    const r = newWs.getRangeByIndexes(0, 0, slice.length, 1);
    r.setValues(slice);
    r.getFormat().autofitColumns();
  }
}
