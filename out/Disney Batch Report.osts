/// <reference path="../office-scripts-docs.d.ts" />

function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();

  const headersBetweenBC = ['PREF Code', 'AKAs', 'ISRCs', 'Setup Note']; // C..F after insert
  const headersBeforeA = ['Assigned To']; // A after insert
  const headersBetweenVW = ['IPI', 'PRO']; // W..X after insert
  const headerFillColor = '#3bc1ff';

  const lookupIsrcSheetName = 'ISRCs for upload';
  const lookupAkaSheetName = 'AKA Table';
  const ipTableSheetName = 'IP Table';
  const lookupIsrcColumn = 'B';
  const lookupAkaColumn = 'A';
  const ipKeyCol = 'B'; // song code
  const ipIpiCol = 'D'; // IPI to return
  const ipProCol = 'E'; // PRO to return

  const styleHeader = (range: ExcelScript.Range) => {
    const fmt = range.getFormat();
    fmt.getFill().setColor(headerFillColor);
    const font = fmt.getFont();
    font.setBold(true);
    font.setUnderline(ExcelScript.RangeUnderlineStyle.single);
  };

  const wrapAndFormat = (range: ExcelScript.Range) => {
    const fmt = range.getFormat();
    fmt.setWrapText(true);
    range.getEntireColumn().getFormat().autofitColumns();
    range.getEntireRow().getFormat().autofitRows();
  };

  const applyNonEmptyFill = (
    columnIndex: number,
    dataRowCount: number,
    color: string
  ) => {
    if (dataRowCount <= 0) return;

    const range = sheet.getRangeByIndexes(1, columnIndex, dataRowCount, 1);
    const topLeftA1 = range.getCell(0, 0).getAddress().split('!')[1]; // e.g. "E2"
    const colLetter = topLeftA1.replace(/\d+/g, ''); // "E"

    const cf = range.addConditionalFormat(
      ExcelScript.ConditionalFormatType.custom
    );
    const custom = cf.getCustom();

    if (custom) {
      custom.getRule().setFormula(`=LEN($${colLetter}2)>0`);
      custom.getFormat().getFill().setColor(color);
    }
  };

  // Build a map from song code -> { ipi, pro } using the IP Table
  const getIpMap = (
    wsName: string,
    keyCol: string,
    ipiCol: string,
    proCol: string
  ): Map<string, { ipi: string; pro: string }> => {
    const ws = workbook.getWorksheet(wsName);
    if (!ws) throw new Error(`Lookup sheet "${wsName}" not found.`);
    const lr = ws.getUsedRange().getRowCount();
    if (lr <= 1) return new Map();

    const keyVals = ws.getRange(`${keyCol}1:${keyCol}${lr}`).getValues(); // [[B1],[B2],...]
    const ipiVals = ws.getRange(`${ipiCol}1:${ipiCol}${lr}`).getValues();
    const proVals = ws.getRange(`${proCol}1:${proCol}${lr}`).getValues();

    const map = new Map<string, { ipi: string; pro: string }>();
    for (let i = 0; i < lr; i++) {
      const rawKey = keyVals[i][0];
      const key =
        rawKey === null || rawKey === undefined
          ? ''
          : (typeof rawKey === 'string' ? rawKey : String(rawKey))
              .trim()
              .toLowerCase();
      if (!key) continue;

      const ipiRaw = ipiVals[i][0];
      const proRaw = proVals[i][0];

      const ipi = ipiRaw == null ? '' : String(ipiRaw).trim();
      const pro = proRaw == null ? '' : String(proRaw).trim();

      // If duplicates exist, later rows will overwrite earlier ones.
      map.set(key, { ipi, pro });
    }
    return map;
  };

  // Delete column G
  sheet.getRange('G:G').delete(ExcelScript.DeleteShiftDirection.left);

  // Insert four columns between B and C
  sheet.getRange('C:F').insert(ExcelScript.InsertShiftDirection.right);
  const newBetweenBCHeaderRange = sheet.getRangeByIndexes(0, 2, 1, 4); // row 1, cols C..F
  newBetweenBCHeaderRange.setValues([headersBetweenBC]);
  styleHeader(newBetweenBCHeaderRange);

  // Insert one column before A
  sheet.getRange('A:A').insert(ExcelScript.InsertShiftDirection.right);
  const newBeforeAHeaderRange = sheet.getRange('A1:A1');
  newBeforeAHeaderRange.setValues([headersBeforeA]);
  styleHeader(newBeforeAHeaderRange);

  // Insert two columns between V and W
  sheet.getRange('W:X').insert(ExcelScript.InsertShiftDirection.right);
  const newBetweenVYHeaderRange = sheet.getRangeByIndexes(0, 22, 1, 2);
  newBetweenVYHeaderRange.setValues([headersBetweenVW]);
  styleHeader(newBetweenVYHeaderRange);

  // Trim whitespace in column B (song codes)
  const used = sheet.getUsedRange();
  const lastRow = used.getRowCount();
  const songCodeColB = sheet.getRangeByIndexes(0, 1, lastRow, 1);
  const trimmedB = songCodeColB
    .getValues()
    .map((row) =>
      row.map((cell) => (typeof cell === 'string' ? cell.trim() : cell))
    );
  songCodeColB.setValues(trimmedB);

  // Build lookup sets (case-insensitive)
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
  let dataRowCount = Math.max(0, lastRow - 1);
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

  if (dataRowCount > 0) {
    const adRange = sheet.getRangeByIndexes(1, 29, dataRowCount, 1); // AD2:AD...
    const aeRange = sheet.getRangeByIndexes(1, 30, dataRowCount, 1); // AE2:AE...
    const gRange = sheet.getRangeByIndexes(1, 6, dataRowCount, 1); // G2:G...

    const adVals = adRange.getValues(); // [[val], ...]
    const aeVals = aeRange.getValues();

    const updatedG: string[][] = Array.from({ length: dataRowCount }, () => [
      '',
    ]);

    for (let i = 0; i < dataRowCount; i++) {
      const adRaw = adVals[i][0];
      const aeRaw = aeVals[i][0];

      const adText =
        adRaw === null || adRaw === undefined
          ? ''
          : (typeof adRaw === 'string'
              ? adRaw.trim()
              : String(adRaw)
            ).toLowerCase();

      const aeHasValue = !(
        aeRaw === null ||
        aeRaw === undefined ||
        aeRaw === ''
      );

      // Regex (case-insensitive) for matches anywhere in the cell
      const isInternational = /(?:international|intl)/i.test(adText);
      const isThemePark = /theme\s*park/i.test(adText);

      let note = '';

      if (aeHasValue) {
        if (isInternational) note = 'foreign FTV';
        else if (isThemePark) note = 'Parks & Resort';
        else note = 'FTV';
      } else {
        if (isThemePark) note = 'Parks & Resort';
        else note = 'Pop Work';
      }

      updatedG[i][0] = note;
    }

    gRange.setValues(updatedG);

    // === IPI (W) and PRO (X) lookups from "IP Table" keyed by song code (B) ===
    // Columns: W=22 (IPI), X=23 (PRO)
    const ipMap = getIpMap(ipTableSheetName, ipKeyCol, ipIpiCol, ipProCol);

    const updatedW: string[][] = Array.from({ length: dataRowCount }, () => [
      '',
    ]); // IPI
    const updatedX: string[][] = Array.from({ length: dataRowCount }, () => [
      '',
    ]); // PRO

    for (let i = 0; i < dataRowCount; i++) {
      const raw = colB_data[i][0];
      const key =
        raw === null || raw === undefined
          ? ''
          : (typeof raw === 'string' ? raw : String(raw)).trim().toLowerCase();

      if (key && ipMap.has(key)) {
        const { ipi, pro } = ipMap.get(key)!;
        updatedW[i][0] = ipi || '';
        updatedX[i][0] = pro || '';
      } else {
        updatedW[i][0] = '';
        updatedX[i][0] = '';
      }
    }

    // Write IPI and PRO
    sheet.getRangeByIndexes(1, 22, dataRowCount, 1).setValues(updatedW); // W2:...
    sheet.getRangeByIndexes(1, 23, dataRowCount, 1).setValues(updatedX); // X2:...

    // Color fill non-empty cells
    applyNonEmptyFill(22, dataRowCount, headerFillColor); // W (IPI)
    applyNonEmptyFill(23, dataRowCount, headerFillColor); // X (PRO)
    applyNonEmptyFill(4, dataRowCount, headerFillColor); // AKA (E)
    applyNonEmptyFill(5, dataRowCount, headerFillColor); // ISRC (F)
    applyNonEmptyFill(6, dataRowCount, headerFillColor); // Setup Note (G)

    // Text-wrap and auto fit headers
    const usedCols = sheet.getUsedRange().getColumnCount();
    const allHeaders = sheet.getRangeByIndexes(0, 0, 1, usedCols);
    wrapAndFormat(allHeaders);
  }
}
