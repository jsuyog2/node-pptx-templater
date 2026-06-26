const { resolve } = require('path');
const JSZip = require('jszip');
const fs = require('fs');

async function main() {
  const originalFile = resolve(__dirname, '../templates/sample.pptx');
  const data = fs.readFileSync(originalFile);
  const zip = await JSZip.loadAsync(data);
  
  const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('text');
  console.log('--- Original ppt/_rels/presentation.xml.rels ---');
  console.log(presRelsXml);
}

main().catch(err => console.error(err));
