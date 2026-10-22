const fs = require('fs');

const files = [
  'js/docx-importer.js',
  'js/parser.js',
  'js/shuffler.js',
  'js/docx-builder.js',
  'js/answer-key.js',
  'js/ui.js',
  'js/app.js'
];

let bundle = '(function() {\n"use strict";\n\n';

for (const f of files) {
  let content = fs.readFileSync(f, 'utf8');
  // Strip import statements (single or multi-line)
  content = content.replace(/import\s+[\s\S]*?from\s*['"][^'"]+['"];?/g, '');
  content = content.replace(/import\s*['"][^'"]+['"];?/g, '');
  // Strip exports
  content = content.replace(/^export\s+(async\s+function|function|const|let|var)\s+/gm, '$1 ');
  content = content.replace(/^export\s*\{[\s\S]*?\};?\s*$/gm, '');
  content = content.replace(/^export\s+default\s+.*?;?\s*$/gm, '');

  bundle += `// --- FILE: ${f} ---\n` + content.trim() + '\n\n';
}

bundle += '\n})();\n';

fs.writeFileSync('js/app.bundle.js', bundle, 'utf8');
console.log('Built js/app.bundle.js successfully! Length:', bundle.length);
