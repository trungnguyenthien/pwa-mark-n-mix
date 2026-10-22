const fs = require('fs');

function bundleFiles(files, outputFile) {
  let bundle = '(function() {\n"use strict";\n\n';

  for (const f of files) {
    let content = fs.readFileSync(f, 'utf8');
    // Strip import statements
    content = content.replace(/import\s+[\s\S]*?from\s*['"][^'"]+['"];?/g, '');
    content = content.replace(/import\s*['"][^'"]+['"];?/g, '');
    // Strip exports
    content = content.replace(/^export\s+(async\s+function|function|const|let|var)\s+/gm, '$1 ');
    content = content.replace(/^export\s*\{[\s\S]*?\};?\s*$/gm, '');
    content = content.replace(/^export\s+default\s+.*?;?\s*$/gm, '');

    bundle += `// --- FILE: ${f} ---\n` + content.trim() + '\n\n';
  }

  bundle += '\n})();\n';

  fs.writeFileSync(outputFile, bundle, 'utf8');
  console.log(`Built ${outputFile} successfully! Length: ${bundle.length}`);
}

// 1. Build app.bundle.js cho trang Trộn Đề (mix.html)
const appFiles = [
  'js/docx-importer.js',
  'js/parser.js',
  'js/shuffler.js',
  'js/docx-builder.js',
  'js/answer-key.js',
  'js/ui.js',
  'js/app.js'
];
bundleFiles(appFiles, 'js/app.bundle.js');

// 2. Build convert.bundle.js cho trang Chuyển Đổi Marker (convert.html)
const convertFiles = [
  'js/docx-importer.js',
  'js/marker-converter.js',
  'js/convert.js'
];
bundleFiles(convertFiles, 'js/convert.bundle.js');
