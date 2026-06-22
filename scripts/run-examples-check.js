const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const examplesDir = path.join(__dirname, '../examples');
const files = fs.readdirSync(examplesDir).filter(f => f.endsWith('.js'));

console.log(`Running ${files.length} examples...`);
let failed = 0;

for (const file of files) {
  const filePath = path.join(examplesDir, file);
  console.log(`\n▶ Running ${file}...`);
  try {
    execSync(`node "${filePath}"`, { stdio: 'inherit', cwd: path.join(__dirname, '..') });
    console.log(`✅ ${file} succeeded`);
  } catch (err) {
    console.error(`❌ ${file} failed:`, err.message);
    failed++;
  }
}

if (failed > 0) {
  console.error(`\n❌ Validation failed: ${failed} example(s) failed`);
  process.exit(1);
} else {
  console.log('\n✅ All examples succeeded!');
}
