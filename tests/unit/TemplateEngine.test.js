/**
 * @fileoverview Unit tests for TemplateEngine.
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { TemplateEngine } from '../../src/core/TemplateEngine.js';
import { XMLParser } from '../../src/parsers/XMLParser.js';

describe('TemplateEngine', () => {
  let engine;

  beforeEach(() => {
    engine = new TemplateEngine(new XMLParser());
  });

  describe('replaceTextInXml()', () => {
    it('should replace a simple placeholder in a text run', () => {
      const slideXml = `<p:sld>
        <a:p><a:r><a:t>{{name}}</a:t></a:r></a:p>
      </p:sld>`;

      const result = engine.replaceTextInXml(slideXml, { '{{name}}': 'John Doe' });
      expect(result).toContain('John Doe');
      expect(result).not.toContain('{{name}}');
    });

    it('should replace multiple placeholders', () => {
      const slideXml = `<p:sld>
        <a:p><a:r><a:t>{{title}}</a:t></a:r></a:p>
        <a:p><a:r><a:t>{{year}}</a:t></a:r></a:p>
      </p:sld>`;

      const result = engine.replaceTextInXml(slideXml, {
        '{{title}}': 'Annual Report',
        '{{year}}': '2026',
      });

      expect(result).toContain('Annual Report');
      expect(result).toContain('2026');
      expect(result).not.toContain('{{title}}');
      expect(result).not.toContain('{{year}}');
    });

    it('should handle fragmented placeholders split across runs', () => {
      // PowerPoint sometimes splits text runs mid-placeholder
      const slideXml = `<p:sld>
        <a:p>
          <a:r><a:t>{{ti</a:t></a:r>
          <a:r><a:t>tle}}</a:t></a:r>
        </a:p>
      </p:sld>`;

      const result = engine.replaceTextInXml(slideXml, { '{{title}}': 'My Title' });
      expect(result).toContain('My Title');
      expect(result).not.toContain('{{ti');
      expect(result).not.toContain('tle}}');
    });

    it('should return original XML when no replacements provided', () => {
      const slideXml = '<p:sld><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:sld>';
      const result = engine.replaceTextInXml(slideXml, {});
      expect(result).toBe(slideXml);
    });

    it('should escape XML special characters in replacement values', () => {
      const slideXml = '<p:sld><a:p><a:r><a:t>{{company}}</a:t></a:r></a:p></p:sld>';
      const result = engine.replaceTextInXml(slideXml, { '{{company}}': 'A & B Corp' });
      expect(result).toContain('A &amp; B Corp');
    });

    it('should not modify XML if placeholder not found', () => {
      const slideXml = '<p:sld><a:p><a:r><a:t>Static text</a:t></a:r></a:p></p:sld>';
      const result = engine.replaceTextInXml(slideXml, { '{{missing}}': 'Value' });
      expect(result).toContain('Static text');
    });

    it('should preserve run formatting when merging fragmented runs', () => {
      const slideXml = `<p:sld>
        <a:p>
          <a:r>
            <a:rPr lang="en-US" sz="2400" b="1"/>
            <a:t>{{ti</a:t>
          </a:r>
          <a:r>
            <a:rPr lang="en-US" sz="1800"/>
            <a:t>tle}}</a:t>
          </a:r>
        </a:p>
      </p:sld>`;

      const result = engine.replaceTextInXml(slideXml, { '{{title}}': 'Hello' });
      // The first run's format should be preserved
      expect(result).toContain('sz="2400"');
      expect(result).toContain('Hello');
    });
  });

  describe('extractPlaceholders()', () => {
    it('should extract all unique placeholders', () => {
      const slideXml = `<p:sld>
        <a:t>{{name}}</a:t>
        <a:t>{{date}}</a:t>
        <a:t>{{name}}</a:t>
      </p:sld>`;

      const placeholders = engine.extractPlaceholders(slideXml);
      expect(placeholders).toContain('{{name}}');
      expect(placeholders).toContain('{{date}}');
      expect(placeholders).toHaveLength(2); // Unique only
    });

    it('should return empty array when no placeholders exist', () => {
      const slideXml = '<p:sld><a:t>Static text</a:t></p:sld>';
      const placeholders = engine.extractPlaceholders(slideXml);
      expect(placeholders).toHaveLength(0);
    });
  });

  describe('containsPlaceholders()', () => {
    it('should return true when placeholder is present', () => {
      const result = engine.containsPlaceholders('Hello {{name}}', { '{{name}}': 'World' });
      expect(result).toBe(true);
    });

    it('should return false when no placeholder is present', () => {
      const result = engine.containsPlaceholders('Hello World', { '{{name}}': 'World' });
      expect(result).toBe(false);
    });
  });
});
