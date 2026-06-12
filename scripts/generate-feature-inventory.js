#!/usr/bin/env node
/**
 * Feature Inventory Generator
 *
 * Parses PPTXTemplater.js and generates a structured inventory of all
 * public methods grouped by category. Outputs:
 *   - docs/feature-inventory.json  (machine-readable)
 *   - docs/feature-inventory.md    (human-readable table)
 *
 * Usage:
 *   node scripts/generate-feature-inventory.js
 *
 * Add to CI:
 *   "docs:inventory": "node scripts/generate-feature-inventory.js"
 */

'use strict'

const fs = require('fs')
const path = require('path')

const PPTX_SOURCE = path.join(__dirname, '../src/core/PPTXTemplater.js')
const OUT_JSON = path.join(__dirname, '../docs/feature-inventory.json')
const OUT_MD = path.join(__dirname, '../docs/feature-inventory.md')

// Category mappings — methods not listed here land in 'Utility'
const CATEGORY_MAP = {
  // Core / Load
  load: 'Core', create: 'Core', fromPresentationXml: 'Core', extractPptx: 'Core', buildPptx: 'Core',
  preload: 'Core', cache: 'Core', fromCache: 'Core', clearCache: 'Core',

  // Slide Management
  useSlide: 'Slides', useAllSlides: 'Slides', addSlide: 'Slides', addSlideFrom: 'Slides',
  duplicateSlide: 'Slides', cloneSlide: 'Slides', deleteSlide: 'Slides', removeSlide: 'Slides',
  moveSlide: 'Slides', insertSlide: 'Slides', reorderSlides: 'Slides', tagSlide: 'Slides',
  getSlides: 'Slides', exportSlides: 'Slides', importSlideFrom: 'Slides', importSlides: 'Slides',

  // Tables
  updateTable: 'Tables', getTableRows: 'Tables', addTableRow: 'Tables', removeTableRow: 'Tables',
  insertTableRow: 'Tables', cloneTableRow: 'Tables', updateCell: 'Tables',
  mergeCells: 'Tables', unmergeCells: 'Tables', getMergedCells: 'Tables',
  validateMergeRegion: 'Tables', isMergedCell: 'Tables', getMergeParent: 'Tables',
  getMergeRegion: 'Tables', splitMergedRegion: 'Tables', cloneMergedRegion: 'Tables',
  autoFitTable: 'Tables', resizeTable: 'Tables', getTables: 'Tables',
  addCellShape: 'Cell Shapes', updateCellShape: 'Cell Shapes', removeCellShape: 'Cell Shapes',
  getCellShape: 'Cell Shapes', getCellBounds: 'Cell Shapes', getCellPosition: 'Cell Shapes',

  // Charts
  updateChart: 'Charts', updateChartData: 'Charts', replaceChartSeries: 'Charts',
  updateChartTitle: 'Charts', updateChartCategories: 'Charts', updateDataLabels: 'Charts',
  getDataLabels: 'Charts', validateDataLabels: 'Charts', validateChartLabels: 'Charts',
  validateSeriesNameLabels: 'Charts', getChartLabelPositions: 'Charts',
  getChartBarPositions: 'Charts', getCharts: 'Charts',
  addTextAtPosition: 'Charts', addTextNearChartLabel: 'Charts',

  // Images
  replaceImage: 'Images', addImage: 'Images', removeImage: 'Images', getImages: 'Images',

  // Shapes
  addShape: 'Shapes', updateShape: 'Shapes', removeShape: 'Shapes', getShape: 'Shapes',
  validateShape: 'Shapes', updateShapeText: 'Shapes', updateShapePosition: 'Shapes',
  updateTextBoxPosition: 'Shapes', cloneShape: 'Shapes', deleteShape: 'Shapes', getShapes: 'Shapes',

  // Text
  replaceText: 'Text', replaceTextByTag: 'Text', replaceMultiple: 'Text',
  findText: 'Text', getTextElements: 'Text', updateText: 'Text',
  getList: 'Text', validateList: 'Text',

  // Hyperlinks
  addHyperlink: 'Hyperlinks', addSlideLink: 'Hyperlinks', addImageLink: 'Hyperlinks',
  addShapeLink: 'Hyperlinks', addTextNavigationLink: 'Hyperlinks',
  addShapeNavigationLink: 'Hyperlinks',

  // Z-Order
  bringForward: 'Z-Order', sendBackward: 'Z-Order', bringToFront: 'Z-Order',
  sendToBack: 'Z-Order', setZIndex: 'Z-Order', moveObjectBefore: 'Z-Order',
  moveObjectAfter: 'Z-Order', reorderObjects: 'Z-Order', getObjectOrder: 'Z-Order',
  applyZOrder: 'Z-Order', getTopMostObject: 'Z-Order', getBottomMostObject: 'Z-Order',
  swapObjects: 'Z-Order', sortObjects: 'Z-Order', normalizeZOrder: 'Z-Order',

  // Output
  save: 'Output', saveToFile: 'Output', saveXml: 'Output', saveToFolder: 'Output',
  saveToStream: 'Output', toBuffer: 'Output', toStream: 'Output',

  // Validation & Debug
  validate: 'Validation', repair: 'Validation', validatePresentation: 'Validation',
  validatePresentationXml: 'Validation', validateSlide: 'Validation',
  validateTable: 'Validation', validateArchive: 'Validation',
  validateCharts: 'Validation', repairCharts: 'Validation',
  validateRelationships: 'Validation', debugRelationships: 'Debug',
  inspectSlide: 'Debug', inspectXML: 'Debug', inspectChart: 'Debug',
  inspectChartXML: 'Debug', debugChartRelationships: 'Debug',

  // Performance
  enablePerformanceProfile: 'Performance', getPerformanceMetrics: 'Performance',
  enableDebug: 'Performance', enableDebugZip: 'Performance',

  // Info
  getInfo: 'Info',
}

function extractPublicMethods(source) {
  const lines = source.split('\n')
  const methods = []
  const seen = new Set()

  const patterns = [
    /^\s{2}(?:async\s+)?([a-zA-Z][a-zA-Z0-9_]*)\s*\(/,
    /^\s{2}static\s+(?:async\s+)?([a-zA-Z][a-zA-Z0-9_]*)\s*\(/,
    /^\s{2}get\s+([a-zA-Z][a-zA-Z0-9_]*)\s*\(\)/,
  ]

  const SKIP = new Set(['constructor', 'class', 'module', 'if', 'for', 'while', 'return'])
  const PRIVATE = /^\s{2}#/

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]
    if (PRIVATE.test(line)) continue

    let methodName = null
    for (const p of patterns) {
      const m = p.exec(line)
      if (m) { methodName = m[1]; break }
    }

    if (!methodName || SKIP.has(methodName) || seen.has(methodName)) continue
    seen.add(methodName)

    const isStatic = line.trim().startsWith('static')
    const isAsync = line.includes('async ')
    const isGetter = line.trim().startsWith('get ')

    let hasDoc = false
    let docPreview = ''
    for (let j = i - 1; j >= Math.max(0, i - 20); j--) {
      const prev = lines[j].trim()
      if (prev === '*/') {
        hasDoc = true
        // Try to extract first description line
        for (let k = j - 1; k >= Math.max(0, j - 15); k--) {
          const dl = lines[k].trim().replace(/^\*\s?/, '')
          if (dl && !dl.startsWith('@') && !dl.startsWith('/*') && dl !== '*') {
            docPreview = dl.substring(0, 80)
            break
          }
        }
        break
      }
      if (!prev.startsWith('*') && !prev.startsWith('/*') && prev !== '' && !prev.startsWith('//')) break
    }

    methods.push({
      name: methodName,
      category: CATEGORY_MAP[methodName] || 'Utility',
      isStatic,
      isAsync,
      isGetter,
      hasDoc,
      description: docPreview,
      line: i + 1,
    })
  }

  return methods
}

// ─── Main ────────────────────────────────────────────────────────────────────

const source = fs.readFileSync(PPTX_SOURCE, 'utf8')
const methods = extractPublicMethods(source)

// Group by category
const grouped = {}
for (const m of methods) {
  if (!grouped[m.category]) grouped[m.category] = []
  grouped[m.category].push(m)
}

// Build JSON output
const inventory = {
  generated: new Date().toISOString(),
  totalMethods: methods.length,
  documentedMethods: methods.filter(m => m.hasDoc).length,
  categories: Object.fromEntries(
    Object.entries(grouped).map(([cat, ms]) => [cat, ms.map(m => ({
      name: m.name,
      isStatic: m.isStatic || undefined,
      isAsync: m.isAsync || undefined,
      isGetter: m.isGetter || undefined,
      hasDoc: m.hasDoc,
      description: m.description || undefined,
      line: m.line,
    }))])
  ),
}

fs.mkdirSync(path.dirname(OUT_JSON), { recursive: true })
fs.writeFileSync(OUT_JSON, JSON.stringify(inventory, null, 2))
console.log(`✓ Generated: ${OUT_JSON}`)

// Build Markdown output
const catOrder = [
  'Core', 'Slides', 'Tables', 'Cell Shapes', 'Charts', 'Images', 'Shapes',
  'Text', 'Hyperlinks', 'Z-Order', 'Output', 'Validation', 'Debug',
  'Performance', 'Info', 'Utility',
]

let md = `# Feature Inventory — node-pptx-templater\n\n`
md += `> Generated: ${new Date().toISOString()}\n\n`
md += `**Total public methods:** ${methods.length}  \n`
md += `**Documented:** ${methods.filter(m => m.hasDoc).length}  \n\n`

for (const cat of catOrder) {
  const ms = grouped[cat]
  if (!ms || ms.length === 0) continue
  md += `## ${cat} (${ms.length})\n\n`
  md += `| Method | Static | Async | Documented | Description |\n`
  md += `|--------|--------|-------|------------|-------------|\n`
  for (const m of ms) {
    md += `| \`${m.name}()\` | ${m.isStatic ? '✓' : ''} | ${m.isAsync ? '✓' : ''} | ${m.hasDoc ? '✓' : '❌'} | ${m.description || ''} |\n`
  }
  md += '\n'
}

fs.writeFileSync(OUT_MD, md)
console.log(`✓ Generated: ${OUT_MD}`)
console.log(`\n📊 Summary: ${methods.length} methods, ${methods.filter(m => m.hasDoc).length} documented, ${methods.filter(m => !m.hasDoc).length} undocumented\n`)
