const fs = require('fs');
const path = require('path');

function getLines(filePath) {
  return fs.readFileSync(filePath, 'utf8').split('\n');
}

function main() {
  const origLines = getLines(path.resolve(__dirname, 'slide3_original_comparison.xml'));
  const savedLines = getLines(path.resolve(__dirname, 'slide3_saved_comparison.xml'));
  
  let diff = [];
  
  // Basic line-by-line diff
  const maxLines = Math.max(origLines.length, savedLines.length);
  for (let i = 0; i < maxLines; i++) {
    const orig = origLines[i] || '';
    const saved = savedLines[i] || '';
    
    if (orig.trim() !== saved.trim()) {
      diff.push(`Line ${i + 1}:`);
      diff.push(`- ${orig}`);
      diff.push(`+ ${saved}`);
      diff.push('');
    }
  }
  
  fs.writeFileSync(path.resolve(__dirname, 'slide3_diff.txt'), diff.join('\n'));
  console.log(`Saved ${diff.length / 4} differences to scratch/slide3_diff.txt`);
}

main();
