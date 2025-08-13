function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();

  const headersBetweenBC = ['PREF Code', 'AKAs', 'ISRCs', 'Setup Note']; // C..F after insert
  const headersBeforeA = ['Assigned To']; // A after insert
  const headerFillColor = '#3bc1ff';

  const lookupIsrcSheetName = 'ISRCs for upload';
  const lookupAkaSheetName = 'AKA Table';
  const lookupIsrcColumn = 'B';
  const lookupAkaColumn = 'A';

  const styleHeader = (range: ExcelScript.Range) => {
    const fmt = range.getFormat();
    fmt.getFill().setColor(headerFillColor);
    const font = fmt.getFont();
    font.setBold(true);
    font.setUnderline(ExcelScript.RangeUnderlineStyle.single);
    range.getEntireColumn().getFormat().autofitColumns();
  };

  // 1) Delete column G
  sheet.getRange('G:G').delete(ExcelScript.DeleteShiftDirection.left);

  // 2) Insert four columns between B and C
  sheet.getRange('C:F').insert(ExcelScript.InsertShiftDirection.right);
  const newBetweenBCHeaderRange = sheet.getRangeByIndexes(0, 2, 1, 4); // row 1, cols C..F
  newBetweenBCHeaderRange.setValues([headersBetweenBC]);
  styleHeader(newBetweenBCHeaderRange);

  // 3) Insert one column before A
  sheet.getRange('A:A').insert(ExcelScript.InsertShiftDirection.right);
  const newBeforeAHeaderRange = sheet.getRange('A1:A1');
  newBeforeAHeaderRange.setValues([headersBeforeA]);
  styleHeader(newBeforeAHeaderRange);

  // 4) Trim whitespace in column B (song codes)
  const used = sheet.getUsedRange();
  const lastRow = used.getRowCount();
  const songCodeColB = sheet.getRangeByIndexes(0, 1, lastRow, 1);
  const trimmedB = songCodeColB
    .getValues()
    .map((row) =>
      row.map((cell) => (typeof cell === 'string' ? cell.trim() : cell))
    );
  songCodeColB.setValues(trimmedB);

  // 5) Build lookup sets (case-insensitive)
  const getLookupSet = (wsName: string, colLetter: string): Set<string> => {
    const ws = workbook.getWorksheet(wsName);
    if (!ws) throw new Error(`Lookup sheet "${wsName}" not found.`);

    const lr = ws.getUsedRange().getRowCount();
    const twoD = ws.getRange(`${colLetter}1:${colLetter}${lr}`).getValues();
    return new Set(
      twoD
        .map((r) => r[0])
        .filter((v) => v !== null && v !== undefined && v !== '')
        .map((v) =>
          typeof v === 'string' ? v.trim().toLowerCase() : String(v)
        )
    );
  };

  const isrcSet = getLookupSet(lookupIsrcSheetName, lookupIsrcColumn);
  const akaSet = getLookupSet(lookupAkaSheetName, lookupAkaColumn);

  // Read B (song codes) for data rows only (skip header)
  const dataRowCount = Math.max(0, lastRow - 1);
  if (dataRowCount === 0) return;

  const colB_data = sheet.getRangeByIndexes(1, 1, dataRowCount, 1).getValues();

  // Decide flags: blank if not matched
  const updatedE: string[][] = new Array(dataRowCount); // AKA
  const updatedF: string[][] = new Array(dataRowCount); // ISRC

  for (let i = 0; i < dataRowCount; i++) {
    const raw = colB_data[i][0];
    const key =
      raw === null || raw === undefined
        ? ''
        : typeof raw === 'string'
        ? raw.trim().toLowerCase()
        : String(raw);

    updatedE[i] = [key && akaSet.has(key) ? 'see AKA table' : ''];
    updatedF[i] = [key && isrcSet.has(key) ? 'ISRCs for upload' : ''];
  }

  sheet.getRangeByIndexes(1, 4, dataRowCount, 1).setValues(updatedE); // E2:...
  sheet.getRangeByIndexes(1, 5, dataRowCount, 1).setValues(updatedF); // F2:...
}
