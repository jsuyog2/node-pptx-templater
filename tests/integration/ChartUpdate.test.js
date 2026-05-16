import { describe, it, expect, beforeAll } from 'vitest';
import { ChartWorkbookUpdater } from '../../src/managers/charts/ChartWorkbookUpdater.js';
import { ChartCacheGenerator } from '../../src/managers/charts/ChartCacheGenerator.js';

describe('Chart Update Engine', () => {
  describe('ChartWorkbookUpdater', () => {
    it('should generate correct column letters', () => {
      expect(ChartWorkbookUpdater.getColumnLetter(0)).toBe('A');
      expect(ChartWorkbookUpdater.getColumnLetter(1)).toBe('B');
      expect(ChartWorkbookUpdater.getColumnLetter(25)).toBe('Z');
      expect(ChartWorkbookUpdater.getColumnLetter(26)).toBe('AA');
    });

    it('should generate correct formula ranges', () => {
      expect(ChartWorkbookUpdater.getFormulaRange('Sheet1', 2, 0, 5, 0)).toBe('Sheet1!$A$2:$A$5');
      expect(ChartWorkbookUpdater.getFormulaSingleCell('Sheet1', 1, 1)).toBe('Sheet1!$B$1');
    });
  });

  describe('ChartCacheGenerator', () => {
    it('should generate strCache correctly', () => {
      const cache = ChartCacheGenerator.generateStrCache(['A', 'B', 'C']);
      expect(cache).toContain('<c:ptCount val="3"/>');
      expect(cache).toContain('<c:pt idx="0"><c:v>A</c:v></c:pt>');
      expect(cache).toContain('<c:pt idx="2"><c:v>C</c:v></c:pt>');
    });

    it('should generate numCache correctly', () => {
      const cache = ChartCacheGenerator.generateNumCache([10, 20.5, 30]);
      expect(cache).toContain('<c:ptCount val="3"/>');
      expect(cache).toContain('<c:formatCode>General</c:formatCode>');
      expect(cache).toContain('<c:pt idx="0"><c:v>10</c:v></c:pt>');
      expect(cache).toContain('<c:pt idx="1"><c:v>20.5</c:v></c:pt>');
    });

    it('should update categories and formulas in xml', () => {
      const xml = '<c:cat><c:strRef><c:f>OldFormula</c:f><c:strCache>OldData</c:strCache></c:strRef></c:cat>';
      const updated = ChartCacheGenerator.updateCategories(xml, ['Cat1', 'Cat2']);
      expect(updated).toContain('<c:f>Sheet1!$A$2:$A$3</c:f>');
      expect(updated).toContain('<c:v>Cat1</c:v>');
      expect(updated).toContain('<c:v>Cat2</c:v>');
    });

    it('should clone dynamic series if requested', () => {
      const xml = '<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Series1</c:v></c:tx></c:ser>';
      const updated = ChartCacheGenerator.appendDynamicSeries(xml, 2);
      expect(updated).toContain('<c:idx val="0"/>');
      expect(updated).toContain('<c:idx val="1"/>');
      expect(updated).toContain('<c:order val="1"/>');
      // Total count of <c:ser> should be 2, but we just check presence
    });
  });
});
