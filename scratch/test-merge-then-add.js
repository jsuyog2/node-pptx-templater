const { PPTXTemplater } = require('../src/index.js');
const path = require('path');
const fs = require('fs');

async function main() {
  const templatePath = path.resolve(__dirname, '../templates/sample.pptx');
  const outputPath = path.resolve(__dirname, 'test-merge-then-add-output.pptx');
  
  console.log('Loading template...');
  const ppt = await PPTXTemplater.load(templatePath);
  ppt.useSlide(3);
  
  console.log('Updating table (first)...');
  ppt.updateTable('Table', [
    ['Name', '', 'Role', 'Dept'],
    ['Alice', '', 'Engineer', 'Platform'],
    ['Bob', '', 'Designer', 'Product']
  ]);
  
  console.log('Merging Bob row columns 0 and 1...');
  ppt.mergeCells('Table', 2, 0, 2, 1);
  
  console.log('Adding nested row...');
  ppt.addTableRow('Table', ['Bob', '', '     Designer', ['Product', 'Value']]);
  
  console.log('Saving to:', outputPath);
  await ppt.saveToFile(outputPath);
  
  // Test opening it in PowerPoint
  console.log('Testing opening in PowerPoint via COM...');
  const execSync = require('child_process').execSync;
  try {
    const resolvedPath = path.resolve(outputPath);
    execSync(`powershell -Command "$ppt = New-Object -ComObject PowerPoint.Application; $ppt.DisplayAlerts = 2; $pres = $ppt.Presentations.Open('${resolvedPath}', 1, 0, 0); $pres.Close(); $ppt.Quit()"`, { stdio: 'inherit' });
    console.log('✅ Success: Opened cleanly without repair errors.');
  } catch (err) {
    console.error('❌ Error opening:', err.message);
  }
}

main().catch(console.error);
