const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function main() {
  const samplePath = path.resolve('templates/sample.pptx');
  const zip = await JSZip.loadAsync(fs.readFileSync(samplePath));
  const xml = await zip.files['ppt/charts/chart3.xml'].async('text');

  const serPattern = /<c:ser>([\s\S]*?)<\/c:ser>/g;
  let serMatch;
  let serIndex = 0;
  while ((serMatch = serPattern.exec(xml)) !== null) {
    const serContent = serMatch[1];
    
    // Find tx
    const txMatch = /<c:tx>([\s\S]*?)<\/c:tx>/.exec(serContent);
    const txVal = txMatch ? txMatch[1].substring(0, 200) : 'None';
    
    console.log(`\n=================== Series #${serIndex++} ===================`);
    console.log(`tx: ${txVal.trim()}`);
    
    // Find dLbls
    const dLblsMatch = /<c:dLbls>([\s\S]*?)<\/c:dLbls>/.exec(serContent);
    if (dLblsMatch) {
      console.log(`dLbls Content:`);
      console.log(dLblsMatch[1]);
    } else {
      console.log(`dLbls Content: None`);
    }
  }
}

main().catch(console.error);
