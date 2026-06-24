const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname);
const srcDir = path.join(root, 'src');
const sourceFile = path.join(root, 'monitor-tsi.user.js');
const outputFile = path.join(root, 'monitor-tsi.user.js');

function slugify(text) {
  return text
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .replace(/-+/g, '-');
}

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function splitSections(lines) {
  const sections = [];
  let current = { name: 'bootstrap', lines: [] };
  const sectionRegex = /^\s*\/\/\s*[\-─]{2,}\s*(.+?)\s*[\-─]{2,}\s*$/;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const match = sectionRegex.exec(line);
    if (match) {
      const title = match[1].trim();
      if (current.lines.length > 0) {
        sections.push(current);
      }
      current = { name: title, lines: [line] };
      continue;
    }
    current.lines.push(line);
  }
  if (current.lines.length > 0) sections.push(current);
  return sections;
}

function writeModuleFiles(sections) {
  ensureDir(srcDir);
  const names = [];
  sections.forEach((section, idx) => {
    const filename = `${String(idx + 1).padStart(2, '0')}-${slugify(section.name) || 'section'}.js`;
    const filePath = path.join(srcDir, filename);
    const content = `// ${section.name}\n${section.lines.join('\n')}\n`;
    fs.writeFileSync(filePath, content, 'utf8');
    names.push(filename);
  });
  return names;
}

function buildOutput(header, moduleFiles) {
  const prefix = `(function() {\n  'use strict';\n\n`;
  const suffix = `\n})();\n`;
  const modules = moduleFiles.map(file => {
    const raw = fs.readFileSync(path.join(srcDir, file), 'utf8');
    const lines = raw.split(/\r?\n/);
    // Remove leading comment from module file if present
    if (lines.length > 0 && lines[0].startsWith('// ')) lines.shift();
    return lines.join('\n').trimEnd() + '\n';
  }).join('\n');
  const final = header.join('\n') + '\n\n' + prefix + modules + suffix;
  fs.writeFileSync(outputFile, final, 'utf8');
}

function main() {
  const text = fs.readFileSync(sourceFile, 'utf8');
  const lines = text.split(/\r?\n/);

  const headerStart = lines.findIndex(line => line.trim() === '// ==UserScript==');
  const headerEnd = lines.findIndex((line, idx) => idx > headerStart && line.trim() === '// ==/UserScript==');
  if (headerStart === -1 || headerEnd === -1) {
    console.error('Could not find userscript metadata header.');
    process.exit(1);
  }
  const header = lines.slice(headerStart, headerEnd + 1);

  const bodyLines = lines.slice(headerEnd + 1);
  // Drop any opening IIFE and closing wrapper from source file if present
  let startIndex = 0;
  while (startIndex < bodyLines.length && bodyLines[startIndex].trim() === '') startIndex++;
  if (bodyLines[startIndex].match(/^\(?function\(|^\(?\(\)\s*=>\s*\{/)) {
    startIndex++;
    if (bodyLines[startIndex] && bodyLines[startIndex].trim() === "'use strict';") startIndex++;
    if (bodyLines[startIndex] && bodyLines[startIndex].trim() === '"use strict";') startIndex++;
  }
  let endIndex = bodyLines.length;
  while (endIndex > startIndex && bodyLines[endIndex - 1].trim() === '') endIndex--;
  if (bodyLines[endIndex - 1].trim() === '})();' || bodyLines[endIndex - 1].trim() === '})();') {
    endIndex--;
  }
  const bodyToSplit = bodyLines.slice(startIndex, endIndex);

  const sections = splitSections(bodyToSplit);
  const moduleFiles = writeModuleFiles(sections);
  buildOutput(header, moduleFiles);
  console.log(`Wrote ${moduleFiles.length} modules to src/ and rebuilt ${path.basename(outputFile)}.`);
}

main();
