const { PPTXTemplater } = require('../src/index.js');
const { resolve } = require('path');
const JSZip = require('jszip');
const fs = require('fs');

async function main() {
  const file = resolve(__dirname, '../tests/fixtures/sample.pptx');
  const ppt = await PPTXTemplater.load(file);
  
  console.log('Original Slide Count:', ppt.slideCount);
  
  // Duplicate slide 2 to position 1
  ppt.duplicateSlide(2, 1);
  
  console.log('New Slide Count:', ppt.slideCount);
  
  const buffer = await ppt.toBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const presXml = await zip.file('ppt/presentation.xml').async('text');
  
  console.log('--- presentation.xml ---');
  console.log(presXml);
}

main().catch(err => console.error(err));
