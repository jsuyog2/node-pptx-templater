const { resolve } = require('path');
const JSZip = require('jszip');
const fs = require('fs');

async function main() {
  const outputFile = resolve(__dirname, '../examples/output/basic-output.pptx');
  const data = fs.readFileSync(outputFile);
  const zip = await JSZip.loadAsync(data);
  const contentTypes = await zip.file('[Content_Types].xml').async('text');
  console.log(contentTypes);
}

main().catch(err => console.error(err));
