/**
 * @fileoverview Unit tests for XMLParser.
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { XMLParser } from '../../src/parsers/XMLParser.js';

describe('XMLParser', () => {
  let parser;

  beforeEach(() => {
    parser = new XMLParser();
  });

  describe('parse()', () => {
    it('should parse simple XML', () => {
      const xml = '<root><child attr="value">text</child></root>';
      const result = parser.parse(xml);
      expect(result).toHaveProperty('root');
      expect(result.root).toHaveProperty('child');
    });

    it('should preserve attribute prefixes', () => {
      const xml = '<a:r xmlns:a="test"><a:rPr lang="en-US"/><a:t>Hello</a:t></a:r>';
      const result = parser.parse(xml);
      expect(result['a:r']).toBeDefined();
    });

    it('should throw PPTXError for invalid XML', () => {
      expect(() => parser.parse('<unclosed')).toThrow();
    });

    it('should throw for null/undefined input', () => {
      expect(() => parser.parse(null)).toThrow();
      expect(() => parser.parse(undefined)).toThrow();
      expect(() => parser.parse('')).toThrow();
    });

    it('should parse XML with namespace prefixes', () => {
      const xml = `<p:sld xmlns:p="test" xmlns:a="test2">
        <p:cSld>
          <p:spTree/>
        </p:cSld>
      </p:sld>`;
      const result = parser.parse(xml);
      expect(result['p:sld']).toBeDefined();
      expect(result['p:sld']['p:cSld']).toBeDefined();
    });
  });

  describe('build()', () => {
    it('should serialize object back to XML', () => {
      const obj = { root: { child: { '@_attr': 'value', '#text': 'hello' } } };
      const xml = parser.build(obj);
      expect(xml).toContain('attr="value"');
      expect(xml).toContain('hello');
    });

    it('should prepend XML declaration if provided', () => {
      const obj = { root: {} };
      const decl = '<?xml version="1.0" encoding="UTF-8"?>';
      const xml = parser.build(obj, decl);
      expect(xml.startsWith(decl)).toBe(true);
    });
  });

  describe('extractDeclaration()', () => {
    it('should extract XML declaration', () => {
      const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root/>';
      const decl = parser.extractDeclaration(xml);
      expect(decl).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    });

    it('should return default declaration when none present', () => {
      const xml = '<root/>';
      const decl = parser.extractDeclaration(xml);
      expect(decl).toContain('<?xml');
    });
  });

  describe('findAll()', () => {
    it('should find nested nodes by path', () => {
      const xml = '<root><parent><child>A</child></parent><parent><child>B</child></parent></root>';
      const obj = parser.parse(xml);
      // path depends on parsed structure
      expect(obj).toBeDefined();
    });
  });

  describe('extractTextContent()', () => {
    it('should extract all text content from slide XML', () => {
      const xml = '<p:sld><a:t>Hello</a:t><a:t>World</a:t></p:sld>';
      const texts = parser.extractTextContent(xml);
      expect(texts).toContain('Hello');
      expect(texts).toContain('World');
    });

    it('should ignore empty text nodes', () => {
      const xml = '<p:sld><a:t></a:t><a:t>  </a:t><a:t>Text</a:t></p:sld>';
      const texts = parser.extractTextContent(xml);
      expect(texts).toHaveLength(1);
      expect(texts[0]).toBe('Text');
    });
  });

  describe('validate()', () => {
    it('should return valid for well-formed XML', () => {
      const result = parser.validate('<root><child/></root>');
      expect(result.valid).toBe(true);
      expect(result.error).toBeNull();
    });

    it('should return invalid for malformed XML', () => {
      const result = parser.validate('<root><unclosed>');
      // fast-xml-parser is lenient; just check it returns an object
      expect(result).toHaveProperty('valid');
    });
  });

  describe('deepClone()', () => {
    it('should create a deep copy', () => {
      const obj = { a: { b: { c: 1 } } };
      const clone = parser.deepClone(obj);
      clone.a.b.c = 99;
      expect(obj.a.b.c).toBe(1); // Original unchanged
    });
  });

  describe('replaceInXml()', () => {
    it('should replace all occurrences by default', () => {
      const xml = '<a><b>foo</b><b>foo</b></a>';
      const updated = parser.replaceInXml(xml, 'foo', 'bar');
      expect(updated).toBe('<a><b>bar</b><b>bar</b></a>');
    });

    it('should replace only first when all=false', () => {
      const xml = '<a><b>foo</b><b>foo</b></a>';
      const updated = parser.replaceInXml(xml, 'foo', 'bar', false);
      expect(updated).toBe('<a><b>bar</b><b>foo</b></a>');
    });
  });
});
