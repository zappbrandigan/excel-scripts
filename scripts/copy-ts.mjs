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

async function main() {
  try {
    // flat match: only files directly under src/
    const files = await glob(path.join(srcDir, '*.ts'));
    await fs.ensureDir(outDir);

    for (const file of files) {
      const fileName = path.basename(toOsts(file));
      const destPath = path.join(outDir, fileName);
      await fs.copy(file, destPath);
      console.log(
        `Copied: ${path.relative(__dirname, file)} → ${path.relative(
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
