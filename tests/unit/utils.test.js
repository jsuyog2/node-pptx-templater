/**
 * @fileoverview Unit tests for utility functions.
 */

import { describe, it, expect } from 'vitest'
import {
  generateRelationshipId,
  parseRelationshipId,
  isValidRelationshipId,
  remapRelationshipIds,
} from '../../src/utils/relationshipUtils.js'
import { validateXML, repairXML, extractAttributeValues } from '../../src/utils/xmlUtils.js'
import { generateUniqueId, generateGuid, generateSlideId } from '../../src/utils/idUtils.js'

describe('relationshipUtils', () => {
  describe('generateRelationshipId()', () => {
    it('should return rId1 for empty array', () => {
      expect(generateRelationshipId([])).toBe('rId1')
    })

    it('should return next sequential ID', () => {
      expect(generateRelationshipId(['rId1', 'rId2', 'rId3'])).toBe('rId4')
    })

    it('should handle gaps in IDs', () => {
      expect(generateRelationshipId(['rId1', 'rId5'])).toBe('rId6')
    })

    it('should handle null/undefined', () => {
      expect(generateRelationshipId(null)).toBe('rId1')
      expect(generateRelationshipId(undefined)).toBe('rId1')
    })
  })

  describe('parseRelationshipId()', () => {
    it('should parse valid rId', () => {
      expect(parseRelationshipId('rId5')).toBe(5)
      expect(parseRelationshipId('rId100')).toBe(100)
    })

    it('should return -1 for invalid IDs', () => {
      expect(parseRelationshipId('foo')).toBe(-1)
      expect(parseRelationshipId('')).toBe(-1)
    })
  })

  describe('isValidRelationshipId()', () => {
    it('should validate correct format', () => {
      expect(isValidRelationshipId('rId1')).toBe(true)
      expect(isValidRelationshipId('rId999')).toBe(true)
    })

    it('should reject invalid format', () => {
      expect(isValidRelationshipId('rId')).toBe(false) // No number
      expect(isValidRelationshipId('rid1')).toBe(false) // Wrong case
      expect(isValidRelationshipId('1')).toBe(false)
    })
  })

  describe('remapRelationshipIds()', () => {
    it('should remap rId references in XML', () => {
      const xml = '<root r:id="rId1" r:embed="rId2"/>'
      const idMap = new Map([
        ['rId1', 'rId5'],
        ['rId2', 'rId6'],
      ])
      const result = remapRelationshipIds(xml, idMap)
      expect(result).toContain('"rId5"')
      expect(result).toContain('"rId6"')
      expect(result).not.toContain('"rId1"')
      expect(result).not.toContain('"rId2"')
    })
  })
})

describe('xmlUtils', () => {
  describe('validateXML()', () => {
    it('should validate well-formed XML', () => {
      const result = validateXML('<root><child/></root>')
      expect(result.valid).toBe(true)
    })

    it('should handle various XML structures', () => {
      const result = validateXML('<?xml version="1.0"?><root attr="value">text</root>')
      expect(result).toHaveProperty('valid')
    })
  })

  describe('repairXML()', () => {
    it('should remove invalid control characters', () => {
      const xml = '<root>\x00\x01text</root>'
      const { xml: repaired, repaired: wasRepaired } = repairXML(xml)
      expect(repaired).not.toContain('\x00')
      expect(repaired).not.toContain('\x01')
      expect(wasRepaired).toBe(true)
    })

    it('should escape unescaped ampersands', () => {
      const xml = '<root>A & B</root>'
      const { xml: repaired, changes } = repairXML(xml)
      expect(repaired).toContain('&amp;')
      expect(changes).toContain('Escaped unescaped ampersands')
    })

    it('should return unchanged XML when already valid', () => {
      const xml = '<?xml version="1.0"?><root><child/></root>'
      const { repaired } = repairXML(xml)
      expect(repaired).toBe(false)
    })
  })

  describe('extractAttributeValues()', () => {
    it('should extract all values for a given attribute', () => {
      const xml = '<a r:id="rId1"/><b r:id="rId2"/><c r:id="rId3"/>'
      const values = extractAttributeValues(xml, 'r:id')
      expect(values).toContain('rId1')
      expect(values).toContain('rId2')
      expect(values).toContain('rId3')
      expect(values).toHaveLength(3)
    })
  })
})

describe('idUtils', () => {
  describe('generateUniqueId()', () => {
    it('should return 1 for empty array', () => {
      expect(generateUniqueId([])).toBe(1)
    })

    it('should return max + 1', () => {
      expect(generateUniqueId([1, 3, 5])).toBe(6)
    })
  })

  describe('generateGuid()', () => {
    it('should generate a GUID in correct format', () => {
      const guid = generateGuid()
      expect(guid).toMatch(
        /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/i
      )
    })

    it('should generate unique GUIDs', () => {
      const guid1 = generateGuid()
      const guid2 = generateGuid()
      expect(guid1).not.toBe(guid2)
    })
  })

  describe('generateSlideId()', () => {
    it('should return 256 for empty array', () => {
      expect(generateSlideId([])).toBe('256')
    })

    it('should return max + 1', () => {
      expect(generateSlideId(['256', '257', '258'])).toBe('259')
    })
  })
})
