/// <reference path="../office-scripts-docs.d.ts" />

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

  // Defensive check (API should always return a Custom object for the given type)
  if (!custom) return;

  custom.getRule().setFormula(formula);
  custom.getFormat().getFill().setColor(color);
}

/**
 * Convert a 1-based column index to A1 letters (1->A, 2->B, 27->AA).
 *
 *  */
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
 *  - style?: string              Built-in table style name (e.g., "TableStyleMedium2"). Default: "TableStyleMedium2".
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
    style = 'TableStyleMedium2',
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

  if (style) table.setPredefinedTableStyle(style);
  if (autofit) table.getRange().getFormat().autofitColumns();

  return table;
}
