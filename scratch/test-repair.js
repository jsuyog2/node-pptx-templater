const { PPTXTemplater } = require('../src/index.js');
const path = require('path');
const fs = require('fs');

async function main() {
  const templatePath = path.resolve(__dirname, '../templates/sample.pptx');
  const outputPath = path.resolve(__dirname, 'test-repair-output.pptx');
  
  console.log('Loading template...');
  const ppt = await PPTXTemplater.load(templatePath);
  ppt.useSlide(3);
  
  console.log('Adding nested table row...');
  ppt.addTableRow('Table', ['Bob', '', '     Designer', ['Product', 'Value']]);
  
  console.log('Saving to:', outputPath);
  await ppt.saveToFile(outputPath);
  
  // Extract and print the resulting table XML
  const AdmZip = require('jszip');
  const data = fs.readFileSync(outputPath);
  const zip = await AdmZip.loadAsync(data);
  const slideXml = await zip.file('ppt/slides/slide3.xml').async('text');
  
  let formattedXml = slideXml
    .replace(/(<a:tr[^>]*>)/g, '\n  $1')
    .replace(/(<\/a:tr>)/g, '$1\n')
    .replace(/(<a:tc[^>]*>)/g, '\n    $1')
    .replace(/(<\/a:tc>)/g, '$1');
  
  fs.writeFileSync(path.resolve(__dirname, 'test_repair_slide3.xml'), formattedXml);
  console.log('XML saved to scratch/test_repair_slide3.xml');
  
  const validation = ppt.validate();
  console.log('Validation result:', validation);
}

main().catch(console.error);
