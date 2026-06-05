const JSZip = require('jszip');
const fs = require('fs');

async function main() {
  const zip = await JSZip.loadAsync(fs.readFileSync('templates/sample.pptx'));
  const chartFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/charts/chart') && k.endsWith('.xml'));

  for (const f of chartFiles) {
    const xml = await zip.files[f].async('text');
    console.log(`\n${'='.repeat(60)}\n=== ${f} ===\n${'='.repeat(60)}`);

    // Is it a stacked bar?
    const isStackedBar = xml.includes('<c:grouping val="stacked"') || xml.includes('<c:grouping val="percentStacked"');
    const isBar = xml.includes('<c:barChart>');
    console.log(`  Bar chart: ${isBar}, Stacked: ${isStackedBar}`);

    // Print all <c:ser> blocks with their dLbls content
    const serPattern = /<c:ser>([\s\S]*?)<\/c:ser>/g;
    let serMatch;
    let serIdx = 0;
    while ((serMatch = serPattern.exec(xml)) !== null) {
      const serContent = serMatch[1];
      const idxMatch = /<c:idx val="(\d+)"/.exec(serContent);
      const nameMatch = /<c:v>([^<]*)<\/c:v>/.exec(serContent);
      console.log(`\n--- Series ${serIdx} (idx=${idxMatch?.[1]}, name=${nameMatch?.[1]}) ---`);

      // Print full dLbls block
      const dLblsMatch = /<c:dLbls>([\s\S]*?)<\/c:dLbls>/.exec(serContent);
      if (dLblsMatch) {
        console.log('\n[c:dLbls block]:');
        console.log(dLblsMatch[0]);
      } else {
        console.log('  No c:dLbls in this series.');
      }
      serIdx++;
    }

    // Print chart-level dLbls if any
    const chartLevelDLbls = /<c:barChart>([\s\S]*?)<\/c:barChart>/.exec(xml);
    if (chartLevelDLbls) {
      const chartDlbls = /<c:dLbls>([\s\S]*?)<\/c:dLbls>/.exec(chartLevelDLbls[1]);
      if (chartDlbls) {
        // Only if it's NOT inside a c:ser
        const serCheck = /<c:ser>/.exec(chartLevelDLbls[1]);
        if (!serCheck || chartLevelDLbls[1].indexOf(chartDlbls[0]) < chartLevelDLbls[1].indexOf('<c:ser>')) {
          console.log('\n[Chart-level c:dLbls]:');
          console.log(chartDlbls[0]);
        }
      }
    }
  }
}

main().catch(console.error);
