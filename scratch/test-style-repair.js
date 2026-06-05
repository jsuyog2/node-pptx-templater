const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const modulePath = '../src/index.js';
const { PPTXTemplater } = require(modulePath);

async function main() {
  const testFile = path.resolve('templates/sample.pptx');
  if (!fs.existsSync(testFile)) {
    console.log("Template file templates/sample.pptx does not exist!");
    return;
  }
  
  const ppt = await PPTXTemplater.load(testFile);
  ppt.useSlide(1).updateChart('Chart', {
    categories: ['Sales'],
    seriesNameLabels: {
      enabled: true,
      position: 'left',
      autoFit: true,
    },
    series: [
      { name: 'Product A', values: [100] },
      { name: 'Product B', values: [150] },
    ],
  });

  const buffer = await ppt.toBuffer();
  fs.writeFileSync('scratch/output.pptx', buffer);
  console.log("Wrote scratch/output.pptx");

  const zip = await JSZip.loadAsync(buffer);
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('text');
  
  // Find generated shapes
  const spPattern = /<p:sp>([\s\S]*?)<\/p:sp>/g;
  let match;
  while ((match = spPattern.exec(slideXml)) !== null) {
    if (match[1].includes('SeriesNameLabel')) {
      console.log("\n=================== GENERATED SHAPE XML ===================");
      console.log(match[0]);
    }
  }
}

main().catch(console.error);
