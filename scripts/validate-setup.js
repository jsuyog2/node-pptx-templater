const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

console.log('🔍 Validating Node PPTX Templater setup...');

const rootDir = path.resolve(__dirname, '..');

// 1. Verify existence of critical entry points
const requiredFiles = [
  'src/index.js',
  'src/index.mjs',
  'src/index.d.ts',
  'src/cli/index.js',
  'README.md',
  'LICENSE',
  'package.json'
];

for (const file of requiredFiles) {
  const filePath = path.join(rootDir, file);
  if (!fs.existsSync(filePath)) {
    console.error(`❌ Missing critical file: ${file}`);
    process.exit(1);
  }
}
console.log('✅ Critical files exist.');

// 2. Validate package.json exports syntax
const pkg = require('../package.json');
if (!pkg.exports || !pkg.exports['.'] || !pkg.exports['.'].require || !pkg.exports['.'].import || !pkg.exports['.'].types) {
  console.error('❌ package.json "exports" configuration is incomplete or incorrect.');
  process.exit(1);
}
console.log('✅ package.json exports are correctly configured.');

// 3. Test CommonJS importing
try {
  const cjsModule = require('../src/index.js');
  if (!cjsModule.PPTXTemplater) {
    throw new Error('PPTXTemplater export missing in CommonJS bundle');
  }
  console.log('✅ CommonJS import validation succeeded.');
} catch (err) {
  console.error('❌ CommonJS import validation failed:', err.message);
  process.exit(1);
}

// 4. Test ESM importing (dynamic import)
async function validateESM() {
  try {
    const esmModule = await import('../src/index.mjs');
    if (!esmModule.PPTXTemplater) {
      throw new Error('PPTXTemplater export missing in ESM bundle');
    }
    console.log('✅ ESM import validation succeeded.');
  } catch (err) {
    console.error('❌ ESM import validation failed:', err.message);
    process.exit(1);
  }
}

// 5. Run other checks
async function main() {
  await validateESM();

  try {
    console.log('\n--- Running Linter ---');
    execSync('npm run lint', { stdio: 'inherit', cwd: rootDir });

    console.log('\n--- Running Unit & Integration Tests ---');
    execSync('npm run test', { stdio: 'inherit', cwd: rootDir });

    console.log('\n--- Running Examples Validation ---');
    execSync('node scripts/run-examples-check.js', { stdio: 'inherit', cwd: rootDir });

    console.log('\n--- Running Docs Validation ---');
    execSync('npm run docs:validate', { stdio: 'inherit', cwd: rootDir });

    console.log('\n🎉 All validations passed successfully! The package is ready.');
  } catch (err) {
    console.error('\n❌ Execution of checks failed:', err.message);
    process.exit(1);
  }
}

main();
