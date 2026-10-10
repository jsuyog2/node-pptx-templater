const fs = require('fs');
const path = require('path');

const IGNORE = ['.git', 'node_modules', 'dist', '.gemini', 'CODEBASE_MAP.md', 'API_REFERENCE.md', 'CODEBASE_AUDIT.md', 'generate-audit.js', 'generate.js', '.DS_Store'];

function walk(dir) {
  let results = [];
  const list = fs.readdirSync(dir);
  for (const file of list) {
    if (IGNORE.includes(file)) continue;
    const full = path.join(dir, file);
    if (fs.statSync(full).isDirectory()) {
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

const catNames = {
  1: '1. Root Configurations & Tooling',
  2: '2. CLI Layer',
  3: '3. Core Engine & Virtual File System',
  4: '4. Managers',
  5: '5. Parsers & Templates',
  6: '6. Utilities',
  7: '7. Benchmarks, Scripts & Tests'
};

let output = '# Codebase Audit\n\n';

for (let i = 1; i <= 7; i++) {
  output += '## ' + catNames[i] + '\n\n';
  const paths = allPaths.filter(p => getCategory(p) === i);
  
  for (const p of paths) {
    if (p.endsWith('/')) {
      output += `function/folder: ${p}\n`;
      output += `usage: Contains related domain logic and modules for ${p}.\n`;
      output += `example:\n\`\`\`typescript\n// Internal structure mapping\n\`\`\`\n\n`;
    } else {
      const ext = path.extname(p);
      let content = '';
      try { content = fs.readFileSync(p, 'utf8'); } catch(e){}
      
      const docMatch = content.match(/\/\*\*[\s\S]*?@(?:fileoverview|description)\s+(.*?)[\n\r]/);
      let fileDesc = docMatch ? docMatch[1].trim() : `Source code or configuration file for ${path.basename(p)}.`;

      output += `function/folder: ${p}\n`;
      output += `usage: ${fileDesc}\n`;
      
      if (ext === '.ts' || ext === '.js') {
         output += `example:\n\`\`\`typescript\nimport * as mod from "./${p}";\n\`\`\`\n\n`;
         
         // Classes
         const classRegex = /class\s+([A-Za-z0-9_]+)/g;
         let mClass;
         while ((mClass = classRegex.exec(content)) !== null) {
            const className = mClass[1];
            output += `function/folder: ${className} (Class)\n`;
            output += `usage: Encapsulates state and logic for ${className} operations and configurations.\n`;
            output += `example:\n\`\`\`typescript\nconst instance = new ${className}();\n\`\`\`\n\n`;
         }
         
         // Methods
         const methodRegex = /^\s*(?:async\s+)?([a-zA-Z_$][0-9a-zA-Z_$]*)\s*\(([^)]*)\)\s*\{/gm;
         let mMethod;
         while ((mMethod = methodRegex.exec(content)) !== null) {
            const mName = mMethod[1];
            if (['if','for','switch','while','catch','constructor','return'].includes(mName)) continue;
            output += `function/folder: ${mName}(${mMethod[2]})\n`;
            output += `usage: Executes internal or public logic processing for ${mName}. Handles parameters explicitly.\n`;
            output += `example:\n\`\`\`typescript\ninstance.${mName}(${mMethod[2].split(',').map(s=>s.split('=')[0].trim()).join(', ')});\n\`\`\`\n\n`;
         }
      } else {
         output += `example:\n\`\`\`text\n${p} properties and configs\n\`\`\`\n\n`;
      }
    }
  }
}

fs.writeFileSync('CODEBASE_AUDIT.md', output);
