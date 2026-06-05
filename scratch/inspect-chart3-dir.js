const JSZip = require('jszip');
const fs = require('fs');

async function main() {
  const zip = await JSZip.loadAsync(fs.readFileSync('templates/sample.pptx'));
  const xml = await zip.files['ppt/charts/chart3.xml'].async('text');
  
  // Find <c:barDir ...>
  const match = /<c:barDir val="([^"]+)"\/>/.exec(xml);
  console.log('barDir match:', match ? match[0] : 'not found');
  
  // Find grouping
  const group = /<c:grouping val="([^"]+)"\/>/.exec(xml);
  console.log('grouping match:', group ? group[0] : 'not found');
}

main().catch(console.error);
