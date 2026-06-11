const fs = require('fs');
const path = require('path');

function getLines(filePath) {
  return fs.readFileSync(filePath, 'utf8').split('\n');
}

function main() {
  const origLines = getLines(path.resolve(__dirname, 'slide3_original_comparison.xml'));
  const savedLines = getLines(path.resolve(__dirname, 'slide3_saved_comparison.xml'));
  
  console.log('Comparing lines...');
  // Let's filter to just the table element parts to make it easier to read
  const origTableLines = origLines.filter(line => line.includes('<a:tbl>') || line.includes('<a:tr>') || line.includes('<a:tc>') || line.includes('rowSpan') || line.includes('vMerge') || line.includes('gridSpan') || line.includes('hMerge'));
  const savedTableLines = savedLines.filter(line => line.includes('<a:tbl>') || line.includes('<a:tr>') || line.includes('<a:tc>') || line.includes('rowSpan') || line.includes('vMerge') || line.includes('gridSpan') || line.includes('hMerge'));
  
  console.log(`Original Table Lines Count: ${origTableLines.length}`);
  console.log(`Saved Table Lines Count: ${savedTableLines.length}`);
  
  console.log('--- ORIGINAL TABLE XML FRAGMENT ---');
  console.log(origTableLines.slice(0, 30).join('\n'));
  
  console.log('--- SAVED TABLE XML FRAGMENT ---');
  console.log(savedTableLines.slice(0, 30).join('\n'));
}

main();
