const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function main() {
  const samplePath = path.resolve('templates/sample.pptx');
  const zip = await JSZip.loadAsync(fs.readFileSync(samplePath));
  const xml = await zip.files['ppt/charts/chart1.xml'].async('text');

  // Let's count categories in chart1.xml
  const catMatch = /<c:cat>([\s\S]*?)<\/c:cat>/.exec(xml);
  if (catMatch) {
    const ptCountMatch = /<c:ptCount val="(\d+)"\/>/.exec(catMatch[1]);
    console.log(`Categories count: ${ptCountMatch ? ptCountMatch[1] : 'Unknown'}`);
  }

  // Let's see the dLbls of first series
  const serPattern = /<c:ser>([\s\S]*?)<\/c:ser>/g;
  let serMatch;
  let serIndex = 0;
  while ((serMatch = serPattern.exec(xml)) !== null) {
    console.log(`\nSeries #${serIndex++}:`);
    const dLblsMatch = /<c:dLbls>([\s\S]*?)<\/c:dLbls>/.exec(serMatch[1]);
    if (dLblsMatch) {
      console.log(dLblsMatch[1]);
    } else {
      console.log('No dLbls');
    }
  }
}

main().catch(console.error);
