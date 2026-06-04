const { PPTXTemplater } = require('../src/index.js');
const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function main() {
  // Simulate a chart title that has:
  // 1. Multi-run (bold word + newline + normal word)
  // 2. Left alignment
  // We'll manually craft the XML and test updateTitle on it

  const { ChartCacheGenerator } = require('../src/managers/charts/ChartCacheGenerator.js');

  // Simulate a chart XML where the title already has:
  // - Left-aligned paragraph
  // - Two paragraphs: first bold, second normal
  const fakeXml = `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="l"/>
            <a:r>
              <a:rPr b="1" sz="1800"/>
              <a:t>Bold Word</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr algn="l"/>
            <a:r>
              <a:rPr b="0" sz="1800"/>
              <a:t>Normal Word</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
      <c:overlay val="0"/>
      <c:spPr><a:noFill/></c:spPr>
      <c:txPr>
        <a:bodyPr rot="0" wrap="square" anchor="ctr"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="l">
            <a:defRPr sz="1800" b="0"/>
          </a:pPr>
          <a:endParaRPr lang="en-US"/>
        </a:p>
      </c:txPr>
    </c:title>
  </c:chart>
</c:chartSpace>`;

  console.log('=== ORIGINAL TITLE ===');
  const orig = /<c:title>([\s\S]*?)<\/c:title>/.exec(fakeXml);
  console.log(orig[0]);

  const updated = ChartCacheGenerator.updateTitle(fakeXml, 'New Bold\nNew Normal');
  console.log('\n=== UPDATED TITLE (with multi-line) ===');
  const upd = /<c:title>([\s\S]*?)<\/c:title>/.exec(updated);
  console.log(upd[0]);

  // Check alignment preserved
  const alignLeft = upd[0].includes('algn="l"');
  const hasBold = upd[0].includes('b="1"');
  const hasNewBoldText = upd[0].includes('New Bold');
  const hasNewNormalText = upd[0].includes('New Normal');
  console.log('\n✅ Results:');
  console.log('  algn="l" preserved:', alignLeft);
  console.log('  bold formatting preserved:', hasBold);
  console.log('  "New Bold" text set:', hasNewBoldText);
  console.log('  "New Normal" text set:', hasNewNormalText);
}

main().catch(console.error);
