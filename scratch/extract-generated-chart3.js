const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function main() {
  const outputPath = path.resolve('examples/output/basic-output.pptx');
  const zip = await JSZip.loadAsync(fs.readFileSync(outputPath));
  const xml = await zip.files['ppt/charts/chart3.xml'].async('text');

  const serPattern = /<c:ser>([\s\S]*?)<\/c:ser>/g;
  let serMatch;
  let serIndex = 0;
  while ((serMatch = serPattern.exec(xml)) !== null) {
    const serContent = serMatch[1];
    
    console.log(`\n=================== Generated Series #${serIndex++} ===================`);
    
    // Find dLbls
    const dLblsMatch = /<c:dLbls>([\s\S]*?)<\/c:dLbls>/.exec(serContent);
    if (dLblsMatch) {
      console.log(dLblsMatch[1]);
    } else {
      console.log(`dLbls Content: None`);
    }
  }
}

main().catch(console.error);
