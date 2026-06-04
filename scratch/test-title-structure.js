const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function main() {
  // Check how PowerPoint renders a chart with BOTH c:tx and c:txPr
  // Look at chart1.xml - it has c:txPr but no c:tx (just a blank title block)
  // Let's find a real PPTX that has a chart title WITH both c:tx text AND c:txPr style

  const samplePath = path.resolve('templates/sample.pptx');
  const zip = await JSZip.loadAsync(fs.readFileSync(samplePath));

  const files = Object.keys(zip.files).filter(k => k.startsWith('ppt/charts/') && k.endsWith('.xml'));
  for (const f of files) {
    const xml = await zip.files[f].async('text');
    const titleMatch = /<c:title>([\s\S]*?)<\/c:title>/.exec(xml);
    if (titleMatch) {
      const titleContent = titleMatch[1];
      const hasTx = titleContent.includes('<c:tx>');
      const hasTxPr = titleContent.includes('<c:txPr>');
      console.log(`${f}: hasTx=${hasTx}, hasTxPr=${hasTxPr}`);
    }
  }

  // Now let's look at the generated output to see how our update looks
  if (fs.existsSync('examples/output/basic-output.pptx')) {
    const outZip = await JSZip.loadAsync(fs.readFileSync('examples/output/basic-output.pptx'));
    const xml = await outZip.files['ppt/charts/chart1.xml'].async('text');
    const titleMatch = /<c:title>([\s\S]*?)<\/c:title>/.exec(xml);
    console.log('\n=== Generated Output Title ===');
    if (titleMatch) console.log(titleMatch[0]);
  }
}

main().catch(console.error);
