/// <reference path="../office-scripts-docs.d.ts" />

/**
 * ---------------------------------------------------------------------------
 * Script Name: DSL Batch Report
 * Author: Brandon Johnson
 * Date: 2025-08-21
 * Version: 1.1.0
 * ---------------------------------------------------------------------------
 * Description:
 * Automates preperation of Disney CWR batch validation report.
 *
 * Features:
 * - Inserts columns and headers for Pref Code, AKAs, ISRCs, Setup Note, Assigned To, IPI, and PRO
 * - Cleans/trims song titles
 * - Performs lookups against "ISRCs for upload", "AKA Table", and "IP Table"
 * - Generates Setup Notes based on AD/AE columns
 * - Applies conditional formatting to highlight non-empty cells
 * - Formats headers with wrap, bold, underline, and fill color
 * - Formats additional sheets as tables and adds filter buttons
 * - Adds "Setup" sheet
 *
 * Notes:
 * - Column layout in main sheet matches the expected pre-insert positions
 *  - The following sheets exist in the workbook:
 *      "ISRCs for upload" (lookup column )
 *      "AKA Table"        (lookup column A)
 *      "IP Table"         (lookup key column A, IPI in D, PRO in E)
 * - Column indexes in this script are 0-based for Office Scripts API
 *
 */
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();

  // -- Config ----------------------------------------------------------------
  const headersBetweenBC = ['PREF Code', 'AKAs', 'ISRCs', 'Setup Note']; // C..F after insert
  const headersBeforeA = ['Assigned To']; // A after insert
  const headersBetweenVW = ['IPI', 'PRO']; // W..X after insert
  const headerFillColor = '#3bc1ff';

  const lookupIsrcSheetName = 'ISRCs for upload';
  const lookupAkaSheetName = 'AKA Table';
  const ipTableSheetName = 'IP Table';
  const lookupIsrcColumn = 'B';
  const lookupAkaColumn = 'A';
  const ipKeyCol = 'A'; // IP Table lookup
  const ipIpiCol = 'D'; // IPI to return
  const ipProCol = 'E'; // PRO to return

  const setupNotes = {
    pop: 'Pop work (unlinked OG/016, check PM1 and US2 for existing work)',
    marketing:
      'Marketing work (unlinked OG/059 with note "In connection with [TITLE NOTE HERE]")',
    parks:
      'Parks & Resorts work (unlinked OG/016 with work note "In connection with [TITLE NOTE HERE] AT [PARK NAME]")',
    dcap: 'DCappella Pop work (unlinked OG/016 with work note "In connection with [TITLE NOTE HERE]")',
    concert:
      'Disney On Classic concert work (unlinked OG/016 with work note "In connection with [TITLE NOTE HERE]")',
    special:
      'Special event work (unlinked OG/016 with work note "In connection with [TITLE NOTE HERE]")',
    podcast:
      'Podcast work (unlinked OG/016 with work note "In connection with [TITLE NOTE HERE]")',
    foreign:
      'Foreign FTV, set up work as ST if no discerning details or cue sheets available',
    hulu: 'Hulu library work (set up as unlinked OG/059 with note "This work is a part of the Hulu Music Library")',
    musical:
      'Musical work (PD type F-THR, works use OG/023), see PD SA8989 for example setup',
    aff: 'PD-affiliated work (OG/059 ST/NYY with note "In connection [TITLE NOTE HERE]")',
    pdbonus:
      'PD bonus content/featurette work (default to CU/053&BI/NYY unless high profile; then use OG/059&ST/NYY)',
    ftv: 'FTV',
    game: 'Video game; Set up PD as F - MUL, CU/053/BI 00:01:00 unless on OST. Add (VIDEO GAME) to end of title. See PD ADL319 for example.',
  };

  // ─--Helpers (local) ----------------------------------------------------
  // Build a list from IP Table rows: { key, ipi, pro }, lowercasing the key.
  function getIpMap(
    wsName: string,
    keyCol: string,
    ipiCol: string,
    proCol: string
  ): { key: string; ipi: string; pro: string }[] {
    const ws = workbook.getWorksheet(wsName);
    if (!ws) throw new Error(`Lookup sheet "${wsName}" not found.`);

    const used = ws.getUsedRange(true);
    const lr = used ? used.getRowCount() : 0;
    if (lr <= 1) return [];

    const keyVals = ws.getRange(`${keyCol}2:${keyCol}${lr}`).getValues(); // skip header
    const ipiVals = ws.getRange(`${ipiCol}2:${ipiCol}${lr}`).getValues();
    const proVals = ws.getRange(`${proCol}2:${proCol}${lr}`).getValues();

    const rows: { key: string; ipi: string; pro: string }[] = [];
    for (let i = 0; i < keyVals.length; i++) {
      const keyRaw = keyVals[i][0];
      const ipiRaw = ipiVals[i][0];
      const proRaw = proVals[i][0];

      const key = keyRaw == null ? '' : String(keyRaw).trim().toLowerCase();
      const ipi = ipiRaw == null ? '' : String(ipiRaw).trim();
      const pro = proRaw == null ? '' : String(proRaw).trim();

      if (key) rows.push({ key, ipi, pro });
    }
    return rows;
  }

  // Case-insensitive set of values from a single column on a sheet.
  function getLookupSet(wsName: string, colLetter: string): Set<string> {
    const ws = workbook.getWorksheet(wsName);
    if (!ws) throw new Error(`Lookup sheet "${wsName}" not found.`);

    const used = ws.getUsedRange(true);
    const lr = used ? used.getRowCount() : 0;
    if (lr < 2) return new Set();

    const twoD = ws.getRange(`${colLetter}2:${colLetter}${lr}`).getValues(); // skip header
    return new Set(
      twoD
        .map((r) => r[0])
        .filter((v) => v !== null && v !== undefined && v !== '')
        .map((v) =>
          typeof v === 'string' ? v.trim().toLowerCase() : String(v)
        )
    );
  }

  // Remove duplicate rows created by ISRCs --------------------------------
  dedupeBySongAndSeq(sheet);

  // -- Column surgery & header setup --------------------------------------
  // Delete column G
  sheet.getRange('G:G').delete(ExcelScript.DeleteShiftDirection.left);

  // Insert four columns between B and C (-> C..F)
  sheet.getRange('C:F').insert(ExcelScript.InsertShiftDirection.right);
  const hdrBetweenBC = sheet.getRangeByIndexes(0, 2, 1, 4);
  hdrBetweenBC.setValues([headersBetweenBC]);
  styleHeader(hdrBetweenBC, { fill: headerFillColor });

  // Insert one column before A
  sheet.getRange('A:A').insert(ExcelScript.InsertShiftDirection.right);
  const hdrBeforeA = sheet.getRange('A1:A1');
  hdrBeforeA.setValues([headersBeforeA]);
  styleHeader(hdrBeforeA, { fill: headerFillColor });

  // Insert two columns between V and W (-> W..X)
  sheet.getRange('W:X').insert(ExcelScript.InsertShiftDirection.right);
  const hdrBetweenVW = sheet.getRangeByIndexes(0, 22, 1, 2);
  hdrBetweenVW.setValues([headersBetweenVW]);
  styleHeader(hdrBetweenVW, { fill: headerFillColor });

  // -- Trim song titles in column C -----------------------------------------
  const usedMain = sheet.getUsedRange(true);
  const lastRow = usedMain ? usedMain.getRowCount() : 0;

  const songTitleColC = sheet.getRangeByIndexes(0, 2, lastRow, 1);
  const trimmedC = songTitleColC
    .getValues()
    .map((row) =>
      row.map((cell) => (typeof cell === 'string' ? cell.trim() : cell))
    );
  songTitleColC.setValues(trimmedC);

  // -- Lookups (case-insensitive) -------------------------------------------
  const isrcSet = getLookupSet(lookupIsrcSheetName, lookupIsrcColumn);
  const akaSet = getLookupSet(lookupAkaSheetName, lookupAkaColumn);

  // Data rows only (skip header)
  const dataRowCount = Math.max(0, lastRow - 1);
  if (dataRowCount === 0) return;

  const colB_data = sheet.getRangeByIndexes(1, 1, dataRowCount, 1).getValues();

  // Flags in E (AKA) and F (ISRC)
  const outE: string[][] = new Array(dataRowCount);
  const outF: string[][] = new Array(dataRowCount);

  for (let i = 0; i < dataRowCount; i++) {
    const raw = colB_data[i][0];
    const key =
      raw == null
        ? ''
        : typeof raw === 'string'
        ? raw.trim().toLowerCase()
        : String(raw);
    outE[i] = [key && akaSet.has(key) ? 'see AKA table' : ''];
    outF[i] = [key && isrcSet.has(key) ? 'ISRCs for upload' : ''];
  }

  sheet.getRangeByIndexes(1, 4, dataRowCount, 1).setValues(outE); // E
  sheet.getRangeByIndexes(1, 5, dataRowCount, 1).setValues(outF); // F

  // -- Setup Notes (G), based on AD/AE ------------------------------------------
  const adRange = sheet.getRangeByIndexes(1, 29, dataRowCount, 1); // AD
  const aeRange = sheet.getRangeByIndexes(1, 30, dataRowCount, 1); // AE
  const gRange = sheet.getRangeByIndexes(1, 6, dataRowCount, 1); // G

  const adVals = adRange.getValues();
  const aeVals = aeRange.getValues();
  const outG: string[][] = Array.from({ length: dataRowCount }, () => ['']);

  for (let i = 0; i < dataRowCount; i++) {
    const adRaw = adVals[i][0];
    const aeRaw = aeVals[i][0];

    const adText =
      adRaw == null
        ? ''
        : (typeof adRaw === 'string'
            ? adRaw.trim()
            : String(adRaw)
          ).toLowerCase();
    const aeHasValue = !(aeRaw === null || aeRaw === undefined || aeRaw === '');
    const adHasValue = !(adRaw === null || adRaw === undefined || adRaw === '');

    const isInternational = /(?:international|intl)/i.test(adText);
    const isThemePark =
      /theme\s*park/i.test(adText) || /disney\s*cruise\s*lines/i.test(adText);
    const isGame = /\(game\)/i.test(adText);

    let note = '';
    if (aeHasValue) {
      if (isInternational) note = setupNotes.foreign;
      else if (isThemePark) note = setupNotes.parks;
      else if (isGame) note = setupNotes.game;
      else note = setupNotes.ftv;
    } else if (adHasValue) {
      if (isThemePark) note = setupNotes.parks;
      else note = setupNotes.pop;
    } else {
      note = '';
    }

    outG[i][0] = note;
  }

  gRange.setValues(outG);

  // -- IPI (W) and PRO (X) via IP Table keyed by IP (T) ------------------------
  const ipRows = getIpMap(ipTableSheetName, ipKeyCol, ipIpiCol, ipProCol);
  const outW: string[][] = Array.from({ length: dataRowCount }, () => ['']); // IPI (W=22)
  const outX: string[][] = Array.from({ length: dataRowCount }, () => ['']); // PRO (X=23)

  const colT_data = sheet.getRangeByIndexes(1, 19, dataRowCount, 1).getValues();

  for (let i = 0; i < dataRowCount; i++) {
    const raw = colT_data[i][0];
    const searchKey = raw == null ? '' : String(raw).trim().toLowerCase();
    if (!searchKey) continue;

    // Find first IP Table row whose key contains the search key
    const found = ipRows.find((row) => row.key.includes(searchKey));
    if (found) {
      outW[i][0] = found.ipi;
      outX[i][0] = found.pro;
    }
  }

  sheet.getRangeByIndexes(1, 22, dataRowCount, 1).setValues(outW); // W
  sheet.getRangeByIndexes(1, 23, dataRowCount, 1).setValues(outX); // X

  // -- Conditional fills for non-empty cells (E, F, W, X)--------------------
  applyNonEmptyFill(sheet, 22, 1, dataRowCount, headerFillColor); // W (IPI)
  applyNonEmptyFill(sheet, 23, 1, dataRowCount, headerFillColor); // X (PRO)
  applyNonEmptyFill(sheet, 4, 1, dataRowCount, headerFillColor); // E (AKA)
  applyNonEmptyFill(sheet, 5, 1, dataRowCount, headerFillColor); // F (ISRC)

  // -- Header row wrapping & auto-fit ---------------------------------------
  const usedCols = (
    sheet.getUsedRange(true) ?? sheet.getRange('A1')
  ).getColumnCount();
  const allHeaders = sheet.getRangeByIndexes(0, 0, 1, usedCols);
  wrapAndFormat(allHeaders);

  // -- Format lookup sheets as tables ---------------------------------------
  formatSheetAsTable(workbook, 'AKA Table', 'AKATable');
  formatSheetAsTable(workbook, 'ISRCs for upload', 'ISRCLookup');
  formatSheetAsTable(workbook, 'IP Table', 'IPTable');

  // -- Build the "Setup" sheet from A, B, C, (empty) Setup Note, H-----------
  buildSetupSheet(workbook, sheet);

  // -- Freeze top row of active sheet
  sheet.getFreezePanes().freezeRows(1);
  sheet.getAutoFilter().apply(sheet.getRangeByIndexes(0, 0, 1, usedCols));
}

/**
 * Style a header range (fill, bold, underline) in one call.
 *
 * By default it applies a light accent fill, bold text, and a single underline.
 * You can override any of these via `options`. Options not provided are left as-is.
 *
 * @param range   The header range to style (e.g., `ws.getRange("A1:Z1")`).
 * @param options Optional overrides:
 *  - fill: Hex/color string for cell background (e.g., "#3bc1ff").
 *  - bold: Whether to bold the font (default: true).
 *  - underline: Whether to apply a single underline (default: true).
 *  - fontColor: Hex/color string for the font color.
 *
 * @example
 * // Basic usage with defaults
 * styleHeader(ws.getRange("A1:Z1"));
 *
 * @example
 * // Custom fill and no underline
 * styleHeader(ws.getRange("A1:D1"), { fill: "#FFD24D", underline: false });
 */
function styleHeader(
  range: ExcelScript.Range,
  options?: {
    fill?: string;
    bold?: boolean;
    underline?: boolean;
    fontColor?: string;
  }
): void {
  const { fill, bold = true, underline = true, fontColor } = options ?? {};

  const fmt = range.getFormat();
  const font = fmt.getFont();

  if (fill) fmt.getFill().setColor(fill);

  if (bold !== undefined) font.setBold(bold);
  if (underline !== undefined) {
    font.setUnderline(
      underline
        ? ExcelScript.RangeUnderlineStyle.single
        : ExcelScript.RangeUnderlineStyle.none
    );
  }
  if (fontColor) font.setColor(fontColor);
}

/**
 * Enable text wrapping for a range and auto-fit its containing rows and columns.
 *
 * This ensures that wrapped text is fully visible without manual resizing.
 * It adjusts the *entire column(s)* and *entire row(s)* intersecting the range.
 *
 * @param range The target range to wrap and auto-fit (e.g., `ws.getRange("B2:D10")`).
 *
 * @example
 * // Wrap and auto-fit just the first column
 * wrapAndFormat(ws.getRange("A1:A100"));
 *
 * @example
 * // Wrap and auto-fit a block of data
 * wrapAndFormat(ws.getRange("B2:F20"));
 */
function wrapAndFormat(range: ExcelScript.Range): void {
  const fmt = range.getFormat();
  fmt.setWrapText(true);

  range.getEntireColumn().getFormat().autofitColumns();
  range.getEntireRow().getFormat().autofitRows();
}

/**
 * Apply a conditional fill to non-empty cells in a single column.
 *
 * Creates a "Custom" conditional format over the specified column slice so that
 * any cell with a non-empty value gets the given background color.
 *
 * @param ws               Target worksheet.
 * @param columnIndex      Zero-based column index (0 = column A).
 * @param startRowIndex    Zero-based start row index for the data region (e.g., 1 to skip header).
 * @param rowCount         Number of rows to include in the rule. No-op if <= 0.
 * @param color            Fill color (e.g., "#FFF2CC" or "yellow").
 * @param clearExisting    Optional. If true, clears existing conditional formats on the target range first. Default: false.
 *
 * @example
 * // Color non-empty cells in column E (index 4) from row 2 (index 1) for 500 rows
 * applyNonEmptyFill(ws, 4, 1, 500, "#E6FFCC");
 */
function applyNonEmptyFill(
  ws: ExcelScript.Worksheet,
  columnIndex: number,
  startRowIndex: number,
  rowCount: number,
  color: string,
  clearExisting = false
): void {
  if (rowCount <= 0) return;

  // Target range: a 1-column slice starting at (startRowIndex, columnIndex)
  const range = ws.getRangeByIndexes(startRowIndex, columnIndex, rowCount, 1);

  if (clearExisting) {
    range.getConditionalFormats().forEach((cf) => cf.delete());
  }

  const firstRow1Based = startRowIndex + 1;
  const colLetter = _toColumnLetters(columnIndex + 1);
  const formula = `=LEN($${colLetter}${firstRow1Based})>0`;

  const cf = range.addConditionalFormat(
    ExcelScript.ConditionalFormatType.custom
  );
  const custom = cf.getCustom();

  if (!custom) return;

  custom.getRule().setFormula(formula);
  custom.getFormat().getFill().setColor(color);
}

/**
 * Convert an entire worksheet's used range into a styled Excel Table.
 *
 * - Uses `getUsedRange(valuesOnly)` (default: true) to avoid trailing formatted blanks.
 * - Optionally deletes any existing table with the same name before creating a new one.
 * - Applies a predefined style and auto-fits columns.
 *
 * @param wb         Target workbook.
 * @param sheetName  Name of the worksheet to format.
 * @param tableName  Name to assign to the new table.
 * @param options
 *  - valuesOnly?: boolean        Use only cells that contain values for the used range. Default: true.
 *  - clearExisting?: boolean     Delete an existing table with the same name (any sheet) before creating. Default: true.
 *  - hasHeaders?: boolean        Treat first row as headers. Default: true.
 *  - style?: string              Built-in table style name (e.g., "TableStyleMedium2"). Default: "TableStyleLight1".
 *  - autofit?: boolean           Auto-fit columns after creation. Default: true.
 *
 * @returns The created table, or `null` if the sheet has no used range.
 *
 * @example
 * const tbl = formatSheetAsTable(wb, "Data", "DataTable", { valuesOnly: true, style: "TableStyleMedium9" });
 */
function formatSheetAsTable(
  wb: ExcelScript.Workbook,
  sheetName: string,
  tableName: string,
  options?: {
    valuesOnly?: boolean;
    clearExisting?: boolean;
    hasHeaders?: boolean;
    style?: string;
    autofit?: boolean;
  }
): ExcelScript.Table | null {
  const {
    valuesOnly = true,
    clearExisting = true,
    hasHeaders = true,
    style = 'TableStyleLight1',
    autofit = true,
  } = options ?? {};

  const ws = wb.getWorksheet(sheetName);
  if (!ws) throw new Error(`Sheet "${sheetName}" not found.`);

  const used = ws.getUsedRange(valuesOnly);
  if (!used) return null;

  if (clearExisting) {
    const dup = wb.getTables().find((t) => t.getName() === tableName);
    if (dup) dup.delete();
  }

  // Add the table on this worksheet, using the address of the used range.
  const table = ws.addTable(used.getAddress(), hasHeaders);
  table.setName(tableName);
  table.setShowFilterButton(true);

  if (style) table.setPredefinedTableStyle(style);
  if (autofit) table.getRange().getFormat().autofitColumns();

  // Freeze first row
  ws.getFreezePanes().freezeRows(1);

  return table;
}

function _toColumnLetters(col1Based: number): string {
  let s = '';
  while (col1Based > 0) {
    col1Based--;
    s = String.fromCharCode(65 + (col1Based % 26)) + s;
    col1Based = Math.floor(col1Based / 26);
  }
  return s;
}

/**
 * Remove duplicate rows from a worksheet based on a composite key:
 * Song Code (column A) + Seq Number (column L).
 *
 * Keeps the first occurrence of each key and deletes later duplicates.
 *
 * @param ws Worksheet to deduplicate.
 */
function dedupeBySongAndSeq(ws: ExcelScript.Worksheet): void {
  const used = ws.getUsedRange(true);
  if (!used || used.getRowCount() <= 1) return;

  // Deduplicate by columns A (index 0) and L (index 11), with a header row.
  used.removeDuplicates([0, 11], true);
}

/**
 * Create/replace a "Setup" sheet from a source sheet:
 * - Columns copied: A, B, C, (new empty) "Setup Note", H
 * - De-duplicates rows by column A (of the new sheet), keeping the first occurrence
 * - Styles header and auto-fits columns
 *
 * @param wb              Target workbook
 * @param source          Source worksheet (main/active sheet)
 * @param headerFillColor Header fill color for styling (e.g., "#3bc1ff")
 * @returns               The created "Setup" worksheet
 */
function buildSetupSheet(
  wb: ExcelScript.Workbook,
  source: ExcelScript.Worksheet
): ExcelScript.Worksheet {
  const name = 'Setup';

  // Recreate sheet fresh
  const existing = wb.getWorksheet(name);
  if (existing) existing.delete();
  const ws = wb.addWorksheet(name);

  // Source bounds
  const used = source.getUsedRange(true);
  const lastRow = used ? used.getRowCount() : 0;
  const dataRowCount = Math.max(0, lastRow - 1);

  // Read source headers (row 1) for A, B, C, H
  const hdrA = source.getRangeByIndexes(0, 0, 1, 1).getValues()[0][0];
  const hdrB = source.getRangeByIndexes(0, 1, 1, 1).getValues()[0][0];
  const hdrC = source.getRangeByIndexes(0, 2, 1, 1).getValues()[0][0];
  const hdrG = source.getRangeByIndexes(0, 6, 1, 1).getValues()[0][0];
  const hdrH = source.getRangeByIndexes(0, 7, 1, 1).getValues()[0][0];

  // Write target headers: A, B, C, Setup Note, H
  ws.getRangeByIndexes(0, 0, 1, 5).setValues([[hdrA, hdrB, hdrC, hdrG, hdrH]]);

  if (dataRowCount > 0) {
    // Read source data rows for A, B, C, H
    const colA = source.getRangeByIndexes(1, 0, dataRowCount, 1).getValues();
    const colB = source.getRangeByIndexes(1, 1, dataRowCount, 1).getValues();
    const colC = source.getRangeByIndexes(1, 2, dataRowCount, 1).getValues();
    const colG = source.getRangeByIndexes(1, 6, dataRowCount, 1).getValues();
    const colH = source.getRangeByIndexes(1, 7, dataRowCount, 1).getValues();

    const out: (string | number | boolean)[][] = new Array(dataRowCount);
    for (let i = 0; i < dataRowCount; i++) {
      out[i] = [colA[i][0], colB[i][0], colC[i][0], colG[i][0], colH[i][0]];
    }
    ws.getRangeByIndexes(1, 0, dataRowCount, 5).setValues(out);

    // Create table over used range (header aware)
    const usedNew = ws.getUsedRange(true);
    if (!usedNew) return ws;
    const table = ws.addTable(usedNew.getAddress(), true);
    table.setName('SetupTable');
    table.setPredefinedTableStyle('TableStyleLight1');
    // Dedupe by first column (B) -> index 1 within the table range
    table.getRange().removeDuplicates([1], true);
    table.getSort().apply([{ key: 4, ascending: true }]);
  }

  // Wrap header & autofit columns for readability
  wrapAndFormat(ws.getRangeByIndexes(0, 0, 1, 5));
  ws.getFreezePanes().freezeRows(1);

  return ws;
}
