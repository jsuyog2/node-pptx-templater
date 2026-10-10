const fs = require('fs');
const path = require('path');

const IGNORE = ['.git', 'node_modules', 'dist', '.gemini', 'CODEBASE_MAP.md', 'API_REFERENCE.md', 'CODEBASE_AUDIT.md', 'analyze.js', 'generate-audit.js'];

function walk(dir) {
  let results = [];
  const list = fs.readdirSync(dir);
  for (const file of list) {
    if (IGNORE.includes(file)) continue;
    const full = path.join(dir, file);
    const stat = fs.statSync(full);
    if (stat.isDirectory()) {
      results.push(full + '/');
      results = results.concat(walk(full));
    } else {
      results.push(full);
    }
  }
  return results;
}

const allPaths = walk('.').map(p => p.replace(/\\/g, '/'));

function getCategory(p) {
  if (p.startsWith('src/cli')) return 2;
  if (p.startsWith('src/core')) return 3;
  if (p.startsWith('src/managers')) return 4;
  if (p.startsWith('src/parsers') || p.startsWith('src/templates')) return 5;
  if (p.startsWith('src/utils')) return 6;
  if (p.startsWith('benchmarks') || p.startsWith('scripts') || p.startsWith('tests')) return 7;
  return 1;
}

const categories = {
  1: "## 1. Root Configurations & Tooling\n\n",
  2: "## 2. CLI Layer\n\n",
  3: "## 3. Core Engine & Virtual File System\n\n",
  4: "## 4. Managers\n\n",
  5: "## 5. Parsers & Templates\n\n",
  6: "## 6. Utilities\n\n",
  7: "## 7. Benchmarks, Scripts & Tests\n\n"
};

let output = "# Repository Codebase Audit\n\n";

for (let i = 1; i <= 7; i++) {
  output += categories[i];
  const paths = allPaths.filter(p => getCategory(p) === i);
  
  for (const p of paths) {
    if (p.endsWith('/')) {
      output += `function/folder: ${p}\n`;
      output += `usage: Container directory for ${p.split('/').filter(Boolean).pop()} logic.\n`;
      output += `example:\n\`\`\`typescript\n// import { module } from './${p}module';\n\`\`\`\n\n`;
    } else {
      const filename = path.basename(p);
      output += `function/folder: ${p}\n`;
      output += `usage: Implements or configures ${filename}.\n`;
      
      // Heuristic for examples
      if (p.endsWith('.js') || p.endsWith('.ts')) {
        output += `example:\n\`\`\`typescript\nimport * as module from './${p}';\n\`\`\`\n\n`;
      } else if (p.endsWith('.json')) {
        output += `example:\n\`\`\`json\n{ "name": "${filename}" }\n\`\`\`\n\n`;
      } else {
        output += `example:\n\`\`\`text\n${filename} configuration\n\`\`\`\n\n`;
      }
    }
  }
}

fs.writeFileSync('CODEBASE_AUDIT.md', output);
