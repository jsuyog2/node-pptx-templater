#!/usr/bin/env node
/**
 * Documentation Validation Script
 *
 * Checks that every public method in PPTXTemplater.js has JSDoc documentation.
 * Exits with code 1 and lists violations if any undocumented methods are found.
 *
 * Usage:
 *   node scripts/validate-docs.js
 *
 * Add to CI:
 *   "docs:validate": "node scripts/validate-docs.js"
 */

'use strict'

const fs = require('fs')
const path = require('path')

const PPTX_SOURCE = path.join(__dirname, '../src/core/PPTXTemplater.js')

// Methods explicitly excluded from doc requirements (intentional aliases or
// utility getters that don't warrant their own JSDoc)
const EXCLUDED = new Set([
  'constructor',
  // Public manager getters — documented via type system
  'zipManager', 'xmlParser', 'contentTypesManager', 'relationshipManager',
  'slideManager', 'chartManager', 'tableManager', 'shapeManager',
  'imageManager', 'textManager', 'hyperlinkManager', 'mediaManager',
  'slideCount',
])

/**
 * Extract all public methods from the PPTXTemplater source file.
 * @param {string} source
 * @returns {Array<{name: string, line: number, hasDoc: boolean}>}
 */
function extractPublicMethods(source) {
  const lines = source.split('\n')
  const methods = []

  // Regex patterns
  const methodPattern = /^\s{2}(?:async\s+)?([a-zA-Z][a-zA-Z0-9_]*)\s*\(/
  const staticMethodPattern = /^\s{2}static\s+(?:async\s+)?([a-zA-Z][a-zA-Z0-9_]*)\s*\(/
  const getterPattern = /^\s{2}get\s+([a-zA-Z][a-zA-Z0-9_]*)\s*\(\)/
  const privatePattern = /^\s{2}#/

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]

    // Skip private methods
    if (privatePattern.test(line)) continue

    let match = methodPattern.exec(line) || staticMethodPattern.exec(line) || getterPattern.exec(line)
    if (!match) continue

    const methodName = match[1]
    if (EXCLUDED.has(methodName)) continue
    if (methodName === 'class' || methodName === 'module') continue

    // Look backwards for JSDoc block
    let hasDoc = false
    for (let j = i - 1; j >= Math.max(0, i - 15); j--) {
      const prev = lines[j].trim()
      if (prev === '*/') {
        hasDoc = true
        break
      }
      if (prev.startsWith('//') || prev === '') continue
      if (!prev.startsWith('*') && !prev.startsWith('/*')) break
    }

    methods.push({ name: methodName, line: i + 1, hasDoc })
  }

  return methods
}

// ─── Main ────────────────────────────────────────────────────────────────────

const source = fs.readFileSync(PPTX_SOURCE, 'utf8')
const methods = extractPublicMethods(source)

const undocumented = methods.filter(m => !m.hasDoc)
const documented = methods.filter(m => m.hasDoc)

console.log(`\n📋 Documentation Validation Report`)
console.log(`   Source: src/core/PPTXTemplater.js`)
console.log(`   Total public methods: ${methods.length}`)
console.log(`   Documented: ${documented.length}`)
console.log(`   Undocumented: ${undocumented.length}`)

if (undocumented.length === 0) {
  console.log('\n✅ All public methods are documented!\n')
  process.exit(0)
} else {
  console.log('\n❌ The following public methods are missing JSDoc:\n')
  undocumented.forEach(m => {
    console.log(`   Line ${String(m.line).padStart(4)}: ${m.name}()`)
  })
  console.log(`\n  Add /** @description ... @param ... @returns */ blocks above each method.`)
  console.log(`  Run "npm run docs:validate" after adding documentation.\n`)
  process.exit(1)
}
