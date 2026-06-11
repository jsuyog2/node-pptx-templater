const fs = require('fs');
const path = require('path');

function normalizeXml(content) {
  return content
    .replace(/<a:bodyPr><\/a:bodyPr>/g, '<a:bodyPr/>')
    .replace(/<a:tcPr><\/a:tcPr>/g, '<a:tcPr/>')
    .replace(/<a:lstStyle><\/a:lstStyle>/g, '<a:lstStyle/>')
    .replace(/<a:endParaRPr><\/a:endParaRPr>/g, '<a:endParaRPr/>')
    .replace(/<a:endParaRPr lang="[^"]+"><\/a:endParaRPr>/g, (m) => m.replace(/><\/a:endParaRPr>/, '/>'))
    .replace(/<a16:rowId([^>]*)\/>/g, '<a16:rowId$1></a16:rowId>') // Normalize rowId formats if needed
    .replace(/ \/>/g, '/>');
}

function main() {
  const orig = normalizeXml(fs.readFileSync(path.resolve(__dirname, 'basic_original_comparison.xml'), 'utf8')).split('\n');
  const saved = normalizeXml(fs.readFileSync(path.resolve(__dirname, 'basic_saved_comparison.xml'), 'utf8')).split('\n');
  
  let diff = [];
  const maxLines = Math.max(orig.length, saved.length);
  for (let i = 0; i < maxLines; i++) {
    const o = (orig[i] || '').trim();
    const s = (saved[i] || '').trim();
    
    // Ignore minor changes in namespaces or self-closing tags
    if (o !== s) {
      // Check if it's just a self-closing tag difference after normalization
      const oNorm = o.replace(/<(\w+:[^>]+)><\/\w+:[^>]+>/g, '<$1/>').replace(/ \/>/g, '/>');
      const sNorm = s.replace(/<(\w+:[^>]+)><\/\w+:[^>]+>/g, '<$1/>').replace(/ \/>/g, '/>');
      
      if (oNorm !== sNorm) {
        diff.push(`Line ${i + 1}:`);
        diff.push(`- ${o}`);
        diff.push(`+ ${s}`);
        diff.push('');
      }
    }
  }
  
  fs.writeFileSync(path.resolve(__dirname, 'basic_diff.txt'), diff.join('\n'));
  console.log(`Saved ${diff.length / 4} semantic differences to scratch/basic_diff.txt`);
}

main();
