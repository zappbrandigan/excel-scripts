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

// Normalize newlines to LF. Do NOT replace with backslash-n; let JSON.stringify
// escape them to "\n" in the final .osts file.
const normalizeNewlines = (text) => text.replace(/\r\n/g, '\n');

// Extract the initial comment block (top-of-file) for README content
function extractInitialCommentBlock(text) {
  const leading = text.replace(/^\s*/, '');

  if (leading.startsWith('/*')) {
    const m = leading.match(/^\/\*[\s\S]*?\*\//);
    if (!m) return null;
    return m[0]
      .replace(/^\/\*+|\*+\/$/g, '')
      .replace(/^\s*\*\s?/gm, '')
      .trim();
  }

  if (leading.startsWith('//')) {
    const m = leading.match(/^(?:\s*\/\/[^\n]*\n?)+/);
    if (!m) return null;
    return m[0].replace(/^\s*\/\/\s?/gm, '').trim();
  }

  return null;
}

function formatReadme(text) {
  const lines = text.split('\n');
  const formatted = lines
    .filter((line) => !/^\s*-{3,}\s*$/.test(line))
    .map((line) => {
      const m = line.match(/^(\s*)([^:\n]+?)\s*:\s*(.*)$/);
      if (!m) return line;
      const [, indent, key, value] = m;
      return `${indent}**${key.trim()}**: ${value}  `;
    });
  return formatted.join('\n').trim();
}

// Extract the nearest comment block above main() for metadata
function extractHeaderMetadata(text) {
  // Find the main function (supports `function main` or `export default function main`)
  const mainIdx = text.search(/\b(?:export\s+default\s+)?function\s+main\s*\(/);
  if (mainIdx === -1) return { version: null, description: null };

  const beforeMain = text.slice(0, mainIdx);

  // Get the last block comment `/* ... */` before main, if any
  const blockComments = [...beforeMain.matchAll(/\/\*[\s\S]*?\*\//g)];
  const lastBlock = blockComments.length
    ? blockComments[blockComments.length - 1][0]
    : null;

  // Or fall back to a run of `//` lines at the end (closest to main)
  // Capture up to ~30 lines back to avoid grabbing too much.
  let lastLineRun = null;
  {
    const tail = beforeMain.slice(Math.max(0, beforeMain.length - 5000)); // safety window
    const m = tail.match(/(?:^|\n)(?:\s*\/\/[^\n]*\n?){1,30}\s*$/);
    lastLineRun = m ? m[0] : null;
  }

  // Prefer block comment if it is closer; otherwise use line-run
  let header = '';
  if (
    lastBlock &&
    (!lastLineRun ||
      beforeMain.lastIndexOf(lastBlock) > beforeMain.lastIndexOf(lastLineRun))
  ) {
    header = lastBlock;
  } else if (lastLineRun) {
    header = lastLineRun;
  }

  if (!header) return { version: null, description: null };

  // Strip comment syntax
  const stripped = header
    .replace(/^\/\*+|\*+\/$/g, '') // remove /* and */
    .replace(/^\s*\*\s?/gm, '') // remove leading * on lines
    .replace(/^\s*\/\/\s?/gm, ''); // remove // prefixes

  // Parse Version
  // Matches: "Version: 1.2.3", "@version: 1.2.3", "version: 1.2.3"
  const vMatch = stripped.match(/(?:@?version|Version)\s*[:=]\s*([^\n*]+)/i);
  const version = vMatch ? vMatch[1].trim() : null;

  // Parse Description: take text after "Description:" (or "@description:")
  // up until next @tag-like label or end. Supports multi-line descriptions.
  const dMatch = stripped.match(
    /(?:@?description|Description)\s*[:=]\s*([\s\S]*?)(?:\n\s*@\w+|\n\s*\w+\s*[:=]|$)/i
  );
  const description = dMatch ? dMatch[1].trim() : null;

  return { version, description };
}

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
      const baseName = path.basename(file, '.ts');
      const fileName = path.basename(toOsts(file));
      const destDir = path.join(outDir, baseName);
      const destPath = path.join(destDir, fileName);

      const srcText = await fs.readFile(file, 'utf8');

      const cleaned = removeOfficeScriptsRef(srcText);
      const readmeText = extractInitialCommentBlock(cleaned);
      const { version: parsedVersion, description: parsedDescription } =
        extractHeaderMetadata(cleaned);

      // Keep real newlines; JSON.stringify will emit "\n" in the .osts.
      const body = normalizeNewlines(cleaned);

      const osts = {
        version: parsedVersion || '0.2.0',
        body,
        description: parsedDescription || '',
        parameterInfo: defaultParameterInfo,
        apiInfo: defaultApiInfo,
      };

      // Compact JSON (Excel is fine with this)
      const json = JSON.stringify(osts);

      await fs.ensureDir(destDir);
      await fs.outputFile(destPath, json, 'utf8');
      if (readmeText) {
        const formattedReadme = formatReadme(readmeText);
        await fs.outputFile(
          path.join(destDir, 'README.md'),
          `${formattedReadme}\n`,
          'utf8'
        );
      }
      console.log(
        `Wrote: ${path.relative(__dirname, file)} → ${path.relative(
          __dirname,
          destPath
        )} (version=${osts.version}${
          osts.description ? `, description="${osts.description}"` : ''
        })`
      );
    }
  } catch (err) {
    console.error('Copy failed:', err);
    process.exit(1);
  }
}

main();
