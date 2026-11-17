/**
 * @fileoverview Generates decoupled JSON data for Next.js documentation portal (docs-site).
 * Targets:
 *   - docs-site/src/data/apiReference.json
 *   - docs-site/src/data/guides.json
 *   - docs-site/src/data/navigation.json
 *   - docs-site/src/data/docsNavigation.json
 *   - docs-site/src/data/features.json
 *   - docs-site/src/data/compatibility.json
 *   - docs-site/src/data/siteMeta.json
 */

const fs = require('fs');
const path = require('path');

const ROOT_DIR = path.resolve(__dirname, '..');
const DOCS_SITE_DATA = path.join(ROOT_DIR, 'docs-site', 'src', 'data');

fs.mkdirSync(DOCS_SITE_DATA, { recursive: true });

// Run features parser
require('./parse-features.js');

console.log('✓ All decoupled data schemas up to date in docs-site/src/data/');
