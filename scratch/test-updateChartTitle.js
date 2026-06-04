const { PPTXTemplater } = require('../src/index.js');
const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function main() {
  const fixturePath = path.resolve('tests/fixtures/sample.pptx');
  const templatePath = path.resolve('templates/sample.pptx');
  const testFile = fs.existsSync(fixturePath) ? fixturePath : templatePath;

  console.log('Using:', testFile);
  const ppt = await PPTXTemplater.load(testFile);
  console.log(`Loaded ${ppt.slideCount} slides`);

  // Use updateChartTitle API
  ppt.useSlide(1).updateChartTitle('Chart', 'My New Styled Title');

  const buffer = await ppt.toBuffer();
  const zip = await JSZip.loadAsync(buffer);

  const chartXml = await zip.files['ppt/charts/chart1.xml'].async('text');
  const titleMatch = /<c:title>([\s\S]*?)<\/c:title>/.exec(chartXml);

  if (titleMatch) {
    console.log('\n=== Title block after updateChartTitle ===');
    console.log(titleMatch[0]);
    
    const hasSpPr = titleMatch[0].includes('<c:spPr>');
    const hasTxPr = titleMatch[0].includes('<c:txPr>');
    const hasNewText = titleMatch[0].includes('My New Styled Title');
    const hasOverlay = titleMatch[0].includes('<c:overlay');
    
    console.log('\n✅ Results:');
    console.log(`  Title text updated: ${hasNewText}`);
    console.log(`  c:spPr preserved: ${hasSpPr}`);
    console.log(`  c:txPr preserved: ${hasTxPr}`);
    console.log(`  c:overlay preserved: ${hasOverlay}`);
  } else {
    console.log('No title found!');
  }
}

main().catch(console.error);
