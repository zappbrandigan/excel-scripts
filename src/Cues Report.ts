/// <reference path="../office-scripts-docs.d.ts" />

/**
 * ---------------------------------------------------------------------------
 * Script Name: Cue Report
 * Author: Brandon Johnson
 * Date: 2025-09-10
 * Version: 1.0.0
 * ---------------------------------------------------------------------------
 * Description:
 * Automates preperation of cue sheet upload report
 *
 * Features:
 * - Adds a new sheet with a summary table
 * - Table contains productions title, counts, and catalog code
 * - Sorts by Catalog, then Production
 * - Shows a Total row summing Count
 * - Writes "Total Cue Sheets" to E2 and the summed total to F2
 *
 * Notes:
 * - Expects cues report file as input
 * - Currently requires user to verify titles for spelling errors/naming convention abbreviations
 *
 */
function main(workbook: ExcelScript.Workbook) {
  const source = workbook.getActiveWorksheet();
  const used = source.getUsedRange(true);
  if (!used) throw new Error('No data found on the active worksheet.');

  const values = used.getValues();
  if (values.length < 2)
    throw new Error('Need at least one data row beneath the header row.');

  // --- Locate columns by header name (case-insensitive) ---
  const header = values[0].map((v) => String(v ?? '').trim());
  const findCol = (names: string[]) => {
    const re = new RegExp(
      `^(${names
        .map((n) => n.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\$&'))
        .join('|')})$`,
      'i'
    );
    return header.findIndex((h) => re.test(h));
  };

  const titleCol = findCol(['Title']);
  const catalogueCol = findCol(['Catalogue', 'Catalog']);
  if (titleCol === -1)
    throw new Error(
      `Couldn't find a "Title" column. Found: ${header.join(', ')}`
    );
  if (catalogueCol === -1)
    throw new Error(
      `Couldn't find a "Catalogue"/"Catalog" column. Found: ${header.join(
        ', '
      )}`
    );

  // --- Helpers ---
  const extractProduction = (rawTitle: string): string => {
    const t = (rawTitle ?? '')
      .toString()
      .replace(/\u00A0/g, ' ')
      .trim();
    if (!t) return '';
    const idx = t.search(/\s{2,}/); // first run of 2+ spaces
    return idx === -1 ? t : t.slice(0, idx).trim();
  };

  const extractCatalogCode = (rawCatalogue: string): string => {
    const c = (rawCatalogue ?? '').toString().trim();
    if (!c) return '';
    const parts = c.split(/\s*-\s*/);
    return (parts[0] ?? '').trim();
  };

  // --- Aggregate: one row per unique Production ---
  const agg = new Map<string, { count: number; catalogs: Set<string> }>();

  for (let r = 1; r < values.length; r++) {
    const titleCell = String(values[r][titleCol] ?? '').trim();
    if (!titleCell) continue;

    const production = extractProduction(titleCell);
    if (!production) continue;

    const catCell = String(values[r][catalogueCol] ?? '').trim();
    const catalogCode = extractCatalogCode(catCell);

    const entry = agg.get(production);
    if (entry) {
      entry.count += 1;
      if (catalogCode) entry.catalogs.add(catalogCode);
    } else {
      const set = new Set<string>();
      if (catalogCode) set.add(catalogCode);
      agg.set(production, { count: 1, catalogs: set });
    }
  }

  // --- Prepare output rows: [Catalog, Production, Count] ---
  const rows: (string | number)[][] = [];
  agg.forEach((v, prod) => {
    const catalogs = Array.from(v.catalogs.values()).sort();
    const catalogCell =
      catalogs.length <= 1 ? catalogs[0] ?? '' : catalogs.join(', ');
    rows.push([catalogCell, prod, v.count]);
  });

  // --- Create/replace the summary sheet ---
  const sheetName = 'Production Summary';
  const existing = workbook.getWorksheet(sheetName);
  if (existing) existing.delete();
  const ws = workbook.addWorksheet(sheetName);

  // Headers
  ws.getRangeByIndexes(0, 0, 1, 3).setValues([
    ['Catalog', 'Production', 'Count'],
  ]);

  // Body
  if (rows.length > 0) {
    ws.getRangeByIndexes(1, 0, rows.length, 3).setValues(rows);
  }

  // Make a table
  const totalRows = Math.max(1, rows.length + 1);
  const table = ws.addTable(ws.getRangeByIndexes(0, 0, totalRows, 3), true);
  table.setName('ProductionSummary');
  table.setPredefinedTableStyle('TableStyleLight10');

  // Ensure Count column is numeric
  table.getColumns()[2].getRangeBetweenHeaderAndTotal().setNumberFormat('0');

  // Sort by Catalog (col 0) then Production (col 1)
  table.getSort().apply([
    { key: 0, ascending: true },
    { key: 1, ascending: true },
  ]);

  // Totals row: turn it on and set SUM of Count via SUBTOTAL
  table.setShowTotals(true);
  const countCol = table.getColumns()[2];
  const colName = countCol.getName(); // should be "Count"
  const totalCell = countCol.getTotalRowRange();
  // Sum the Count column (ignores filtered-out rows)
  totalCell.setFormula(`=SUBTOTAL(109,${table.getName()}[${colName}])`);
  totalCell.setNumberFormat('0');

  // Summary cells
  ws.getRange('E2').setValue('Total Cue Sheets');
  // Link F2 to the totals row cell for the Count column
  ws.getRange('F2').setFormula(`=${table.getName()}[[#Totals],[${colName}]]`);
  ws.getRange('F2').setNumberFormat('0');

  // Clean up
  const usedOut = ws.getUsedRange(true);
  if (usedOut) {
    usedOut.getFormat().autofitColumns();
    usedOut.getFormat().autofitRows();
  }
  ws.activate();
}
