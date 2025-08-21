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
    .replace(/\r?\n/g, '\\n') // real newlines -> literal "\n"
    .replace(/\t/g, '  '); // optional: normalize tabs

// Minimal, broadly compatible shells for parameterInfo/apiInfo
const defaultParameterInfo = JSON.stringify({
  originalParameterOrder: [],
  parameterSchema: {
    type: 'object',
    default: {},
    'x-ms-visibility': 'internal',
  },
  returnSchema: { type: 'object', properties: {} },
});

const defaultApiInfo = JSON.stringify({
  variant: 'synchronous',
  variantVersion: 2,
});

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

      const osts = {
        version: '0.2.0',
        body: oneLiner,
        description: '',
        parameterInfo: defaultParameterInfo,
        apiInfo: defaultApiInfo,
      };

      // Pretty-print for readability; Excel is fine with compact or pretty JSON
      const json = JSON.stringify(osts);

      await fs.outputFile(destPath, json, 'utf8');
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
