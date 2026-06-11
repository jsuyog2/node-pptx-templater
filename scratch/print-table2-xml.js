const AdmZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function extractTable2Xml(pptxPath) {
  const data = fs.readFileSync(pptxPath);
  const zip = await AdmZip.loadAsync(data);
  const slideXml = await zip.file('ppt/slides/slide3.xml').async('text');
  
  const frameRegex = /<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g;
  let match;
  while ((match = frameRegex.exec(slideXml)) !== null) {
    const frame = match[0];
    if (frame.includes('name="Table2"')) {
      const tblMatch = /<a:tbl>[\s\S]*?<\/a:tbl>/.exec(frame);
      if (tblMatch) {
        return tblMatch[0];
      }
    }
  }
  return null;
}

function formatXml(content) {
  if (!content) return '';
  return content
    .replace(/(<a:tr[^>]*>)/g, '\n  $1')
    .replace(/(<\/a:tr>)/g, '$1\n')
    .replace(/(<a:tc[^>]*>)/g, '\n    $1')
    .replace(/(<\/a:tc>)/g, '$1');
}

async function main() {
  const origTbl = await extractTable2Xml(path.resolve(__dirname, 'test-table2-nesting-output.pptx'));
  const savedTbl = await extractTable2Xml(path.resolve(__dirname, 'test-table2-nesting-saved.pptx'));
  
  fs.writeFileSync(path.resolve(__dirname, 'table2_original.xml'), formatXml(origTbl));
  fs.writeFileSync(path.resolve(__dirname, 'table2_saved.xml'), formatXml(savedTbl));
  console.log('Saved Table2 XMLs to scratch/table2_original.xml and scratch/table2_saved.xml');
}

main().catch(console.error);
