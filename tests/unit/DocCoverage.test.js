import { describe, it, expect } from 'vitest'
import fs from 'fs'
import path from 'path'
import { fileURLToPath } from 'url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)
const PPTX_SOURCE = path.resolve(__dirname, '../../src/core/PPTXTemplater.js')

const EXCLUDED = new Set([
  'constructor',
  'zipManager',
  'xmlParser',
  'contentTypesManager',
  'relationshipManager',
  'slideManager',
  'chartManager',
  'tableManager',
  'shapeManager',
  'imageManager',
  'textManager',
  'hyperlinkManager',
  'mediaManager',
  'slideCount',
])

function extractPublicMethods(source) {
  const lines = source.split('\n')
  const methods = []

  const methodPattern = /^\s{2}(?:async\s+)?([a-zA-Z][a-zA-Z0-9_]*)\s*\(/
  const staticMethodPattern = /^\s{2}static\s+(?:async\s+)?([a-zA-Z][a-zA-Z0-9_]*)\s*\(/
  const getterPattern = /^\s{2}get\s+([a-zA-Z][a-zA-Z0-9_]*)\s*\(\)/
  const privatePattern = /^\s{2}#/

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]
    if (privatePattern.test(line)) continue

    const match =
      methodPattern.exec(line) || staticMethodPattern.exec(line) || getterPattern.exec(line)
    if (!match) continue

    const methodName = match[1]
    if (EXCLUDED.has(methodName)) continue
    if (methodName === 'class' || methodName === 'module') continue

    let hasDoc = false
    for (let j = i - 1; j >= 0; j--) {
      const prevLine = lines[j].trim()
      if (prevLine === '' || prevLine.startsWith('//')) continue
      if (prevLine.endsWith('*/')) {
        for (let k = j; k >= 0; k--) {
          if (lines[k].trim().startsWith('/**')) {
            hasDoc = true
            break
          }
        }
      }
      break
    }

    methods.push({ name: methodName, line: i + 1, hasDoc })
  }

  return methods
}

describe('Documentation Coverage for PPTXTemplater', () => {
  it('should have JSDoc for all public methods in PPTXTemplater.js', () => {
    const source = fs.readFileSync(PPTX_SOURCE, 'utf8')
    const methods = extractPublicMethods(source)
    const undocumented = methods.filter(m => !m.hasDoc)

    if (undocumented.length > 0) {
      const list = undocumented.map(m => `  - ${m.name} (line ${m.line})`).join('\n')
      expect.fail(`Found ${undocumented.length} public method(s) without JSDoc:\n${list}`)
    }

    expect(methods.length).toBeGreaterThan(50)
  })
})
