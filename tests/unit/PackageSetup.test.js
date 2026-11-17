import { describe, it, expect } from 'vitest'
import fs from 'fs-extra'
import path from 'path'
import { fileURLToPath } from 'url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)
const rootDir = path.resolve(__dirname, '../..')

describe('Package Setup & Exports Verification', () => {
  it('should have all critical entry points and root files', () => {
    const requiredFiles = [
      'src/index.js',
      'src/index.mjs',
      'src/index.d.ts',
      'src/cli/index.js',
      'README.md',
      'LICENSE',
      'package.json',
      'showcase/examples.js',
    ]

    for (const file of requiredFiles) {
      const filePath = path.join(rootDir, file)
      expect(fs.existsSync(filePath), `Expected ${file} to exist`).toBe(true)
    }
  })

  it('should have valid package.json exports configuration', () => {
    const pkg = fs.readJsonSync(path.join(rootDir, 'package.json'))
    expect(pkg.exports).toBeDefined()
    expect(pkg.exports['.']).toBeDefined()
    expect(pkg.exports['.'].require).toBe('./src/index.js')
    expect(pkg.exports['.'].import).toBe('./src/index.mjs')
    expect(pkg.exports['.'].types).toBe('./src/index.d.ts')
  })

  it('should load successfully via CommonJS', async () => {
    const { createRequire } = await import('module')
    const require = createRequire(import.meta.url)
    const cjs = require('../../src/index.js')
    expect(cjs.PPTXTemplater).toBeDefined()
    expect(typeof cjs.PPTXTemplater.load).toBe('function')
  })

  it('should load successfully via ESM', async () => {
    const esm = await import('../../src/index.mjs')
    expect(esm.PPTXTemplater).toBeDefined()
    expect(typeof esm.PPTXTemplater.load).toBe('function')
  })
})
