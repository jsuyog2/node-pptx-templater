const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function main() {
  const pptxPath = path.resolve('templates/sample.pptx');
  if (!fs.existsSync(pptxPath)) {
    console.log("Template file templates/sample.pptx does not exist!");
    return;
  }
  const buffer = fs.readFileSync(pptxPath);
  const zip = await JSZip.loadAsync(buffer);
  
  if (zip.file('ppt/slides/slide1.xml')) {
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text');
    console.log("Slide 1 XML Length:", slideXml.length);
    // Find all c:chart references
    const chartRIdMatch = /<c:chart[^>]*r:id="([^"]+)"/g;
    let match;
    while ((match = chartRIdMatch.exec(slideXml)) !== null) {
      console.log("Found chart with relationship ID:", match[1]);
    }
  } else {
    console.log("ppt/slides/slide1.xml not found in zip");
  }

  if (zip.file('ppt/charts/chart1.xml')) {
    const chartXml = await zip.file('ppt/charts/chart1.xml').async('text');
    console.log("Chart 1 XML Length:", chartXml.length);
    // Print chart type elements
    const chartTypes = ['c:barChart', 'c:lineChart', 'c:pieChart', 'c:areaChart'];
    for (const ct of chartTypes) {
      if (chartXml.includes(ct)) {
        console.log("Chart 1 type:", ct);
      }
    }
    // Find grouping and barDir
    const grouping = /<c:grouping\s+val="([^"]+)"\/>/.exec(chartXml);
    const barDir = /<c:barDir\s+val="([^"]+)"\/>/.exec(chartXml);
    console.log("Grouping:", grouping ? grouping[1] : "not found");
    console.log("barDir:", barDir ? barDir[1] : "not found");
    
    // Check for valAx or catAx
    const maxVal = /<c:max\s+val="([^"]+)"\/>/.exec(chartXml);
    console.log("Max val axis limit in XML:", maxVal ? maxVal[1] : "none");
  }
}

main().catch(console.error);
