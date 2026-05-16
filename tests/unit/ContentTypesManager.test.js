/**
 * @fileoverview Unit tests for ContentTypesManager.
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { ContentTypesManager } from '../../src/managers/ContentTypesManager.js';
import { XMLParser } from '../../src/parsers/XMLParser.js';

class MockZipManager {
  constructor(files = {}) {
    this.files = files;
  }
  hasFile(path) {
    return !!this.files[path];
  }
  async readFile(path) {
    return this.files[path] || null;
  }
  writeFile(path, content) {
    this.files[path] = content;
  }
}

describe('ContentTypesManager', () => {
  let parser;
  const initialXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`;

  beforeEach(() => {
    parser = new XMLParser();
  });

  it('should initialize and parse existing content types from ZipManager', async () => {
    const mockZip = new MockZipManager({
      '[Content_Types].xml': initialXml
    });
    const manager = new ContentTypesManager(parser);
    await manager.initialize(mockZip);

    // Verify it parses defaults and overrides internally
    manager.addDefault('png', 'image/png');
    manager.addOverride('/ppt/slides/slide2.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');

    manager.flush(mockZip);
    const flushedXml = mockZip.files['[Content_Types].xml'];
    expect(flushedXml).toContain('Extension="png"');
    expect(flushedXml).toContain('PartName="/ppt/slides/slide2.xml"');
  });

  it('should prevent duplicate default extension and overrides registration', async () => {
    const mockZip = new MockZipManager({
      '[Content_Types].xml': initialXml
    });
    const manager = new ContentTypesManager(parser);
    await manager.initialize(mockZip);

    // Register existing extension and override
    manager.addDefault('xml', 'application/xml');
    manager.addOverride('/ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml');

    manager.flush(mockZip);
    const flushedXml = mockZip.files['[Content_Types].xml'];

    // Ensure they appear exactly once
    const xmlMatches = flushedXml.match(/Extension="xml"/g) || [];
    expect(xmlMatches.length).toBe(1);

    const sldMatches = flushedXml.match(/PartName="\/ppt\/presentation\.xml"/g) || [];
    expect(sldMatches.length).toBe(1);
  });

  it('should delete overrides cleanly', async () => {
    const mockZip = new MockZipManager({
      '[Content_Types].xml': initialXml
    });
    const manager = new ContentTypesManager(parser);
    await manager.initialize(mockZip);

    manager.removeOverride('ppt/slides/slide1.xml');
    manager.flush(mockZip);

    const flushedXml = mockZip.files['[Content_Types].xml'];
    expect(flushedXml).not.toContain('PartName="/ppt/slides/slide1.xml"');
    expect(flushedXml).toContain('PartName="/ppt/presentation.xml"');
  });
});
