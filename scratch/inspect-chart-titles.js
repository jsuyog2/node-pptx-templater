const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function main() {
  const samplePath = path.resolve('templates/sample.pptx');
  const zip = await JSZip.loadAsync(fs.readFileSync(samplePath));
  
  // Look at all chart files
  const files = Object.keys(zip.files).filter(k => k.startsWith('ppt/charts/chart') && k.endsWith('.xml'));
  for (const f of files) {
    const xml = await zip.files[f].async('text');
    const titleMatch = /<c:title>([\s\S]*?)<\/c:title>/.exec(xml);
    if (titleMatch) {
      console.log(`\n=== ${f} - Title block ===`);
      console.log(titleMatch[0]);
    } else {
      console.log(`\n=== ${f} - No title ===`);
    }
  }
}

main().catch(console.error);
