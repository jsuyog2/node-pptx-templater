const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function main() {
  // Simulate what updateTitle does now
  const { ChartCacheGenerator } = require('../src/managers/charts/ChartCacheGenerator.js');
  
  const samplePath = path.resolve('templates/sample.pptx');
  const zip = await JSZip.loadAsync(fs.readFileSync(samplePath));
  const xml = await zip.files['ppt/charts/chart1.xml'].async('text');

  console.log('=== ORIGINAL TITLE ===');
  const titleMatch = /<c:title>([\s\S]*?)<\/c:title>/.exec(xml);
  if (titleMatch) console.log(titleMatch[0]);

  const updated = ChartCacheGenerator.updateTitle(xml, 'New Dynamic Chart Title');

  console.log('\n=== UPDATED TITLE ===');
  const updatedTitleMatch = /<c:title>([\s\S]*?)<\/c:title>/.exec(updated);
  if (updatedTitleMatch) console.log(updatedTitleMatch[0]);
}

main().catch(console.error);
