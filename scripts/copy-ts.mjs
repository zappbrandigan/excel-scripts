// scripts/copy-ts.mjs
import { glob } from 'glob';
import fs from 'fs-extra';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const srcDir = path.join(__dirname, '..', 'src');
const outDir = path.join(__dirname, '..', 'out');

function toOsts(filePath) {
  return filePath.replace(/\.ts$/i, '.osts');
}

// Remove any triple-slash reference to office-scripts-docs.d.ts (any path)
const removeOfficeScriptsRef = (text) =>
  text.replace(
    /^\s*\/\/\/\s*<reference\s+path=["'][^"']*office-scripts-docs\.d\.ts["']\s*\/>\s*\r?\n?/gim,
    ''
  );

// Convert to single-line with \n literals
const toSingleLine = (text) =>
  text
    .replace(/\r?\n/g, '\\n') // replace real newlines with literal "\n"
    .replace(/\t/g, '  '); // optional: expand tabs to spaces for safety

async function main() {
  try {
    const files = await glob(path.join(srcDir, '*.ts')); // flat match
    await fs.ensureDir(outDir);

    for (const file of files) {
      const fileName = path.basename(toOsts(file));
      const destPath = path.join(outDir, fileName);

      const srcText = await fs.readFile(file, 'utf8');
      const cleaned = removeOfficeScriptsRef(srcText);
      const oneLiner = toSingleLine(cleaned);

      await fs.outputFile(destPath, oneLiner, 'utf8');
      console.log(
        `Wrote: ${path.relative(__dirname, file)} → ${path.relative(
          __dirname,
          destPath
        )}`
      );
    }
  } catch (err) {
    console.error('Copy failed:', err);
    process.exit(1);
  }
}

main();
