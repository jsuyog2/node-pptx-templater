import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { existsSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import fsExtra from 'fs-extra';
import JSZip from 'jszip';
import { XMLParser } from '../../src/parsers/XMLParser.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const FIXTURES_DIR = resolve(__dirname, '../fixtures');
const OUTPUT_DIR = resolve(__dirname, '../output-corruption');

let PPTXTemplater;

beforeAll(async () => {
  const module = await import('../../src/index.js');
  PPTXTemplater = module.PPTXTemplater;
  await fsExtra.ensureDir(OUTPUT_DIR);
});

afterAll(async () => {
  await fsExtra.remove(OUTPUT_DIR);
});

describe('PPTXTemplater - Slide Management Corruption Prevention', () => {
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx');
  const runFixtureTests = existsSync(fixtureFile);

  if (!runFixtureTests) {
    it.skip('Fixture file not found — skipping slide corruption tests.', () => {});
    return;
  }

  it('should remove a slide, keep metadata synchronized, and clean sections without corruption', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile);

    // We expect the original template to have 3 slides
    expect(ppt.getInfo().slideCount).toBe(3);

    // Remove slide 2
    ppt.removeSlide(2);
    expect(ppt.getInfo().slideCount).toBe(2);

    // Save to a new file
    const outPath = resolve(OUTPUT_DIR, 'slide-corruption-test.pptx');
    await ppt.saveToFile(outPath);

    // Load the saved file as a ZIP to verify structural integrity
    const zip = await JSZip.loadAsync(await fsExtra.readFile(outPath));
    const xmlParser = new XMLParser();

    // 1. Verify presentation.xml
    const presentationXml = await zip.file('ppt/presentation.xml').async('text');
    const presObj = xmlParser.parse(presentationXml);

    // Slide ID 257 (second slide originally) should be removed from sldIdLst
    const sldIds = presObj['p:presentation']['p:sldIdLst']['p:sldId'];
    const sldIdValues = sldIds.map(s => s['@_id']);
    expect(sldIdValues).not.toContain('257');
    expect(sldIdValues).toContain('256');
    expect(sldIdValues).toContain('258');

    // Slide ID 257 should be removed from all sections
    const extLst = presObj['p:presentation']['p:extLst'];
    if (extLst?.['p:ext']) {
      const exts = Array.isArray(extLst['p:ext']) ? extLst['p:ext'] : [extLst['p:ext']];
      for (const ext of exts) {
        const sectionLst = ext['p14:sectionLst'];
        if (sectionLst?.['p14:section']) {
          const sections = sectionLst['p14:section'];
          for (const section of sections) {
            const sectionSldIds = section['p14:sldIdLst']?.['p14:sldId'] || [];
            const ids = sectionSldIds.map(s => s['@_id']);
            expect(ids).not.toContain('257');
          }
        }
      }
    }

    // 2. Verify docProps/app.xml
    const appXml = await zip.file('docProps/app.xml').async('text');
    const appObj = xmlParser.parse(appXml);
    const properties = appObj.Properties;

    // Slides count should be 2
    expect(Number(properties.Slides)).toBe(2);

    // Slide Titles HeadingPair count should be 2
    const variants = properties.HeadingPairs['vt:vector']['vt:variant'];
    let slideTitlesCount = 0;
    for (let i = 0; i < variants.length; i++) {
      if (variants[i]['vt:lpstr'] === 'Slide Titles') {
        slideTitlesCount = parseInt(variants[i + 1]['vt:i4'], 10);
        break;
      }
    }
    expect(slideTitlesCount).toBe(2);

    // TitlesOfParts should have 6 elements and slide titles should match
    const lpstrs = properties.TitlesOfParts['vt:vector']['vt:lpstr'];
    expect(lpstrs.length).toBe(6);
    // Last two elements should be slide titles
    expect(lpstrs[4]).toBe('Hello {{title}}'); // Slide 1
    expect(lpstrs[5]).toBe('Hello {{title}}'); // Slide 3 (formerly slide 3)
  });
});
