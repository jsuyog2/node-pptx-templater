const { resolve } = require('path');
const JSZip = require('jszip');
const fs = require('fs');

async function main() {
  const originalFile = resolve(__dirname, '../templates/sample.pptx');
  const outputFile = resolve(__dirname, '../examples/output/basic-output.pptx');
  
  const zipOrig = await JSZip.loadAsync(fs.readFileSync(originalFile));
  const zipOut = await JSZip.loadAsync(fs.readFileSync(outputFile));
  
  const filesOrig = Object.keys(zipOrig.files).sort();
  const filesOut = Object.keys(zipOut.files).sort();
  
  console.log('--- Files in Original but NOT in Output ---');
  for (const f of filesOrig) {
    if (!filesOut.includes(f)) {
      console.log(f);
    }
  }
  
  console.log('\n--- Files in Output but NOT in Original ---');
  for (const f of filesOut) {
    if (!filesOrig.includes(f)) {
      console.log(f);
    }
  }
}

main().catch(err => console.error(err));
