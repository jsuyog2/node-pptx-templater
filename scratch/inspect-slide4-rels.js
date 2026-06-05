const JSZip = require('jszip');
const fs = require('fs');

async function main() {
  const zip = await JSZip.loadAsync(fs.readFileSync('templates/sample.pptx'));
  
  // Slide 4 rels
  const relsXml = await zip.files['ppt/slides/_rels/slide4.xml.rels'].async('text');
  console.log('Slide 4 relationships:\n', relsXml);
}

main().catch(console.error);
