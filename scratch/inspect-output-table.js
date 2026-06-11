const AdmZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function main() {
  const pptxPath = path.resolve(__dirname, '../examples/output/basic-output.pptx');
  console.log('Reading:', pptxPath);
  if (!fs.existsSync(pptxPath)) {
    console.error('File does not exist:', pptxPath);
    return;
  }
  const data = fs.readFileSync(pptxPath);
  const zip = await AdmZip.loadAsync(data);
  const slideXml = await zip.file('ppt/slides/slide3.xml').async('text');
  
  // Format the XML slightly for better reading (split by tc, tr, tbl)
  let formattedXml = slideXml
    .replace(/(<a:tr[^>]*>)/g, '\n  $1')
    .replace(/(<\/a:tr>)/g, '$1\n')
    .replace(/(<a:tc[^>]*>)/g, '\n    $1')
    .replace(/(<\/a:tc>)/g, '$1');
  
  fs.writeFileSync(path.resolve(__dirname, 'slide3_output.xml'), formattedXml);
  console.log('Saved formatted slide3 XML to scratch/slide3_output.xml');
}

main().catch(console.error);
