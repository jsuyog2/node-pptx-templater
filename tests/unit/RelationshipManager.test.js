/**
 * @fileoverview Unit tests for RelationshipManager.
 */

import { describe, it, expect, beforeEach } from 'vitest'
import { RelationshipManager, REL_TYPES } from '../../src/managers/RelationshipManager.js'
import { XMLParser } from '../../src/parsers/XMLParser.js'

/**
 * Creates a mock ZipManager for testing.
 */
function createMockZipManager(files = {}) {
  const writtenFiles = new Map()

  return {
    listFiles: prefix => Object.keys(files).filter(f => f.startsWith(prefix) && !files[f]?.dir),
    readFile: async path => files[path] || null,
    writeFile: (path, content) => writtenFiles.set(path, content),
    hasFile: path => path in files,
    _writtenFiles: writtenFiles,
  }
}

const SAMPLE_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>`

describe('RelationshipManager', () => {
  let manager
  let xmlParser

  beforeEach(() => {
    xmlParser = new XMLParser()
    manager = new RelationshipManager(xmlParser)
  })

  describe('initialize()', () => {
    it('should discover .rels files from ZIP', async () => {
      const mockZip = createMockZipManager({
        'ppt/slides/_rels/slide1.xml.rels': SAMPLE_RELS_XML,
      })
      await manager.initialize(mockZip)
      const rels = manager.getRelationships('ppt/slides/slide1.xml')
      expect(rels).toHaveLength(3)
    })

    it('should handle empty ZIP gracefully', async () => {
      const mockZip = createMockZipManager({})
      await manager.initialize(mockZip)
      const rels = manager.getRelationships('ppt/slides/slide1.xml')
      expect(rels).toHaveLength(0)
    })
  })

  describe('getRelsPath()', () => {
    it('should compute correct .rels path for a file', () => {
      expect(manager.getRelsPath('ppt/slides/slide1.xml')).toBe('ppt/slides/_rels/slide1.xml.rels')
    })

    it('should handle root-level files', () => {
      expect(manager.getRelsPath('ppt/presentation.xml')).toBe('ppt/_rels/presentation.xml.rels')
    })
  })

  describe('getRelationships()', () => {
    it('should return parsed relationships', async () => {
      const mockZip = createMockZipManager({
        'ppt/slides/_rels/slide1.xml.rels': SAMPLE_RELS_XML,
      })
      await manager.initialize(mockZip)

      const rels = manager.getRelationships('ppt/slides/slide1.xml')
      expect(rels[0]).toEqual({
        id: 'rId1',
        type: REL_TYPES.SLIDE_LAYOUT,
        target: '../slideLayouts/slideLayout1.xml',
        targetMode: null,
      })
    })
  })

  describe('getRelationshipById()', () => {
    it('should return specific relationship by ID', async () => {
      const mockZip = createMockZipManager({
        'ppt/slides/_rels/slide1.xml.rels': SAMPLE_RELS_XML,
      })
      await manager.initialize(mockZip)

      const rel = manager.getRelationshipById('ppt/slides/slide1.xml', 'rId3')
      expect(rel).not.toBeNull()
      expect(rel.target).toBe('https://example.com')
      expect(rel.targetMode).toBe('External')
    })

    it('should return null for non-existent ID', async () => {
      const mockZip = createMockZipManager({})
      await manager.initialize(mockZip)
      const rel = manager.getRelationshipById('ppt/slides/slide1.xml', 'rId99')
      expect(rel).toBeNull()
    })
  })

  describe('addRelationship()', () => {
    it('should add a new relationship and auto-assign rId', async () => {
      const mockZip = createMockZipManager({
        'ppt/slides/_rels/slide1.xml.rels': SAMPLE_RELS_XML,
      })
      await manager.initialize(mockZip)

      const newId = manager.addRelationship(
        'ppt/slides/slide1.xml',
        REL_TYPES.HYPERLINK,
        'https://new.com',
        'External'
      )

      expect(newId).toBe('rId4') // rId1, rId2, rId3 exist → next is rId4
      const rels = manager.getRelationships('ppt/slides/slide1.xml')
      expect(rels).toHaveLength(4)
    })

    it('should create .rels entry for parts with no existing rels', async () => {
      const mockZip = createMockZipManager({})
      await manager.initialize(mockZip)

      const newId = manager.addRelationship(
        'ppt/slides/slide5.xml',
        REL_TYPES.SLIDE_LAYOUT,
        '../slideLayouts/slideLayout1.xml'
      )

      expect(newId).toBe('rId1')
    })
  })

  describe('removeRelationship()', () => {
    it('should remove a relationship', async () => {
      const mockZip = createMockZipManager({
        'ppt/slides/_rels/slide1.xml.rels': SAMPLE_RELS_XML,
      })
      await manager.initialize(mockZip)
      manager.removeRelationship('ppt/slides/slide1.xml', 'rId2')

      const rels = manager.getRelationships('ppt/slides/slide1.xml')
      expect(rels).toHaveLength(2)
      expect(rels.find(r => r.id === 'rId2')).toBeUndefined()
    })
  })

  describe('resolveTarget()', () => {
    it('should resolve relative paths', () => {
      const resolved = manager.resolveTarget(
        'ppt/slides/slide1.xml',
        '../slideLayouts/slideLayout1.xml'
      )
      expect(resolved).toBe('ppt/slideLayouts/slideLayout1.xml')
    })

    it('should return external URLs as-is', () => {
      const resolved = manager.resolveTarget('ppt/slides/slide1.xml', 'https://example.com')
      expect(resolved).toBe('https://example.com')
    })
  })
})
