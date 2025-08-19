/// <reference path="../office-scripts-docs.d.ts" />

/**
 * ---------------------------------------------------------------------------
 * Company: Universal Music Publishing Group
 * Script Name: Distribute Names to Column
 * Author: Brandon Johnson
 * Date: 2025-08-19
 * Version: 1.0.0
 * ---------------------------------------------------------------------------
 * Description:
 * This Office Script distributes names from a predefined list into a specified
 * column of the active worksheet. It skips the header row and applies names
 * only to data rows.
 *
 * Features:
 * - Choose the target column by its letter (e.g., "C" or "AA").
 * - Select a distribution mode:
 *   • "roundRobin" → cycle through names row-by-row (Alice, Bob, Charlie, …).
 *   • "block" → assign names in contiguous blocks, dividing rows evenly.
 * - Names are defined in a configurable array within the script.
 *
 * Parameters:
 * - columnLetter: string
 *   The Excel column letter where names should be applied.
 * - mode: "roundRobin" | "block"
 *   How names are distributed across rows (default: "roundRobin").
 *
 * Usage:
 * - Enter the column letter and mode when prompted.
 * - Script fills the selected column with names according to the chosen mode.
 *
 * Notes:
 * - The header row is assumed to be the first row of the used range.
 * - Adjust the `NAMES` array inside the script to fit your team’s needs.
 */

/**
 * Distributes names from a predefined list into the chosen column.
 * Choose "roundRobin" to cycle names row-by-row or "block" for contiguous groups.
 *
 * @param workbook The active workbook (provided automatically).
 * @param columnLetter The target column letter (e.g., "C", "AA").
 * @param mode Distribution mode: "roundRobin" or "block".
 */
function main(
  workbook: ExcelScript.Workbook,
  columnLetter: string,
  mode: 'roundRobin' | 'block' = 'roundRobin'
): number {
  const sheet = workbook.getActiveWorksheet();

  // Attention!: EDIT THIS LIST to your desired names
  const NAMES = [
    'Gerlitz, Melissa',
    'Kahl, Eileen',
    'Leko, Jay',
    'Lewis, Mikayla',
    'Mitchell, Emori',
    'Newsome, Houston',
    'Norman, Caleb',
  ];

  if (!NAMES.length) {
    throw new Error('NAMES list is empty. Add at least one name.');
  }

  const used = sheet.getUsedRange(true);
  if (!used) {
    console.log('No used range on the sheet.');
    return 0;
  }

  const totalRows = used.getRowCount();
  const dataRows = Math.max(0, totalRows - 1); // exclude header
  if (dataRows === 0) {
    console.log('No data rows found (only header present).');
    return 0;
  }

  const colIndex = columnLetterToIndex(columnLetter);
  const firstUsedRow = used.getRowIndex(); // usually 0
  const firstDataRowIndex = firstUsedRow + 1;

  const targetRange = sheet.getRangeByIndexes(
    firstDataRowIndex,
    colIndex,
    dataRows,
    1
  );

  let values: (string | number | boolean)[][];

  if (mode === 'block') {
    values = buildBlockValues(dataRows, NAMES);
  } else {
    values = buildRoundRobinValues(dataRows, NAMES);
  }

  targetRange.setValues(values);

  console.log(`Mode: ${mode}`);
  console.log(`Total rows (including header): ${totalRows}`);
  console.log(`Data rows (excluding header): ${dataRows}`);
  console.log(
    `Filled column ${columnLetter.toUpperCase()} with ${dataRows} names.`
  );
  return dataRows;
}

/** Convert Excel column letters (e.g., "A", "Z", "AA") to zero-based index. */
function columnLetterToIndex(letter: string): number {
  const s = letter.trim().toUpperCase();
  if (!/^[A-Z]+$/.test(s)) {
    throw new Error(`Invalid column letter: "${letter}"`);
  }
  let index = 0;
  for (let i = 0; i < s.length; i++) {
    index = index * 26 + (s.charCodeAt(i) - 64); // 'A'->1 ... 'Z'->26
  }
  return index - 1; // zero-based
}

/** Round-robin distribution: A, B, C, A, B, C, ... */
function buildRoundRobinValues(
  rowCount: number,
  names: string[]
): (string | number | boolean)[][] {
  const v: (string | number | boolean)[][] = [];
  for (let i = 0; i < rowCount; i++) {
    v.push([names[i % names.length]]);
  }
  return v;
}

/**
 * Block distribution:
 * - Compute near-equal block sizes for each name.
 * - Example: 10 rows, 3 names -> [4, 3, 3]
 * - Fills all rows for first name, then second, etc.
 */
function buildBlockValues(
  rowCount: number,
  names: string[]
): (string | number | boolean)[][] {
  const n = names.length;
  const base = Math.floor(rowCount / n);
  let remainder = rowCount % n;

  const v: (string | number | boolean)[][] = [];
  for (let i = 0; i < n; i++) {
    const take = base + (remainder > 0 ? 1 : 0);
    remainder = Math.max(0, remainder - 1);
    for (let j = 0; j < take; j++) {
      v.push([names[i]]);
      if (v.length === rowCount) return v; // stop early if rows are filled
    }
  }

  // In case there are more rows than computed (shouldn't happen), pad round-robin:
  while (v.length < rowCount) {
    v.push([names[v.length % n]]);
  }
  return v;
}
