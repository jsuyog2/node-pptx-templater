export class PPTXTemplater {
  static load(source: string | Buffer | any, options?: any): Promise<PPTXTemplater>;
  static setLogLevel(level: string): void;
  static preload(source: string | Buffer): Promise<any>;
  static cache(source: string | Buffer): Promise<any>;
  static fromCache(source: string | Buffer): Promise<PPTXTemplater>;
  static clearCache(): void;
  static fromPresentationXml(options: any): Promise<PPTXTemplater>;
  static create(): Promise<PPTXTemplater>;
  static extractPptx(pptxPath: string, outputPath: string, options?: any): Promise<void>;
  static buildPptx(folderPath: string, pptxPath: string): Promise<void>;

  slideCount: number;
  useSlide(slideNumber: number): this;
  setSlideDimensions(width: number, height: number, options?: any): this;
  setSlideSize(width: number, height: number): this;
  setAspectRatio(ratio: '16:9' | '16:10' | '4:3'): this;
  getSlideDimensions(): { width: number; height: number; widthEmu: number; heightEmu: number; type: string };
  getSlideSize(): { width: number; height: number; type?: string };
  getSlideLayout(slideNumber?: number): { name: string; type: string; master: string; rId?: string };
  alignShapes(shapeIds: string[], alignment: 'left' | 'right' | 'center' | 'top' | 'bottom' | 'middle'): this;
  distributeShapes(shapeIds: string[], direction: 'horizontal' | 'vertical'): this;
  checkCollision(idA: string, idB: string): boolean;
  isContained(innerId: string, outerId: string): boolean;
  getElements(filter?: any): any[];
  findElements(filterOrPredicate?: any): any[];
  appendTextRun(textOrOptions: any, runOptions?: any): this;
  prependTextRun(textOrOptions: any, runOptions?: any): this;
  replaceText(replacements: Record<string, string | number>): this;
  updateChart(chartName: string, data: any): this;
  updateChartTitle(chartName: string, title: string): this;
  applyZOrder(slideNumber: number, zOrderConfig: any[]): this;
  removeSlide(slideNumber: number): this;
  addHyperlink(options: any): this;
  getHyperlinks(slideIndex?: number): any[];
  removeHyperlink(textOrOptions: any): this;
  addSlideLink(options: any): this;
  addTable(tableIdOrOptions: any, options?: any): string;
  getTableRows(tableName: string, options?: any): Promise<any[]> | any[];
  updateTable(tableName: string, rows: any[][]): this;
  addTableRow(tableName: string, row: any[], options?: any): this;
  removeTableRow(tableName: string, rowIndex: number): this;
  addColumn(tableId: string, colIndexOrData: any, colDataOrOptions?: any, options?: any): this;
  deleteColumn(tableId: string, colIndex: number): this;
  stripeRows(tableId: string, options?: any): this;
  adjustTableRow(tableName: string, options: any): this;
  adjustTableColumn(tableName: string, options: any): this;
  adjustTable(tableName: string, options?: any): this;
  resize(objectIdOrName: string, options: any): this;
  mergeCells(tableName: string, startRow: number, startCol: number, endRow: number, endCol: number): this;
  splitTableOnOverflow(tableName: string, options?: any): Promise<{ splitCount: number; slidesCreated: number }>;
  saveToFile(filePath: string, options?: any): Promise<void>;
  toFile(filePath: string, options?: any): Promise<void>;
  save(filePath: string, options?: any): Promise<void>;
  saveXml(folderPath: string): Promise<void>;
  saveToFolder(folderPath: string): Promise<void>;
  toBuffer(options?: any): Promise<Buffer>;
  toArrayBuffer(options?: any): Promise<ArrayBuffer>;
  toBase64(options?: { dataUri?: boolean }): Promise<string>;
  toExpress(res: any, filename?: string, options?: any): Promise<void>;
  toStream(options?: any): Promise<any>;
  saveToStream(writableOrOptions: any, options?: any): Promise<void>;
  exportSlides(...slideNumbers: number[]): Promise<any>;
  importSlideFrom(sourceEngine: PPTXTemplater, slideRef: any): Promise<any>;
  repair(): Promise<this>;
  inspectXML(xmlPath: string): Promise<any>;
  validateCharts(): Promise<any>;
  repairCharts(): Promise<any>;
  inspectChartXML(chartFileName: string): Promise<any>;
  getDataLabels(chartId: string, options?: any): Promise<any>;
  validateDataLabels(chartId: string, options?: any): Promise<any>;
  validateChartLabels(chartId: string, options?: any): Promise<any>;
  validateSeriesNameLabels(chartId: string, options?: any): Promise<any>;
  getChartLabelPositions(chartId: string): Promise<any>;
  getChartBarPositions(chartId: string): Promise<any>;
  addShape(typeOrOptions: any, options?: any): Promise<any>;
  updateShape(shapeId: string, options: any): Promise<any>;
  removeShape(shapeId: string): Promise<any>;
  addCellShape(tableId: string, rowIndex: number, colIndex: number, options: any): Promise<any>;
  updateCellShape(tableId: string, rowIndex: number, colIndex: number, shapeIndex: number, options: any): Promise<any>;
  removeCellShape(tableId: string, rowIndex: number, colIndex: number, shapeIndex: number): Promise<any>;
  replaceImage(imageIdOrName: string, sourcePathOrBuffer: string | Buffer): Promise<any>;
  addImage(sourcePathOrBuffer: string | Buffer, options?: any): Promise<any>;
  validatePresentation(): Promise<any>;
  validatePresentationXml(): Promise<any>;
  validateSlide(slideIndex: number): Promise<any>;
  validateTable(tableId: string): Promise<any>;
  validateArchive(): Promise<any>;
}

export const PPTXTemplate: typeof PPTXTemplater;

export class PackageLoader {
  static load(source: any): Promise<VirtualFileSystem>;
}
export function isPkZip(buf: Uint8Array | Buffer): boolean;

export class VirtualFileSystem {
  entries: Map<string, Uint8Array>;
  has(path: string): boolean;
  get(path: string): Uint8Array | undefined;
  set(path: string, data: Uint8Array | Buffer): void;
  delete(path: string): boolean;
  getText(path: string): string | undefined;
  setText(path: string, text: string): void;
}

export class ZipManager {
  [key: string]: any;
}

export class XMLParser {
  [key: string]: any;
}

export const Z_ORDER_SYMBOL: unique symbol;

export class SlideManager {
  [key: string]: any;
}

export class ChartManager {
  [key: string]: any;
}

export class TableManager {
  [key: string]: any;
}

export class ShapeManager {
  [key: string]: any;
}

export class ImageManager {
  [key: string]: any;
}

export class TextManager {
  [key: string]: any;
}

export class HyperlinkManager {
  [key: string]: any;
}

export class MediaManager {
  [key: string]: any;
}

export class RelationshipManager {
  [key: string]: any;
}

export class ZOrderManager {
  [key: string]: any;
}

export class OutputWriter {
  [key: string]: any;
}

export class TemplateEngine {
  [key: string]: any;
}

export class ValidationEngine {
  [key: string]: any;
}

export const EMU_PER_PT: number;
export const EMU_PER_IN: number;
export const EMU_PER_PX: number;
export function ptToEmu(pt: number): number;
export function inToEmu(inches: number): number;
export function pxToEmu(px: number): number;
export function emuToPt(emu: number): number;
export function emuToIn(emu: number): number;
export function emuToPx(emu: number): number;
export function toEmu(value: string | number | undefined | null): number;
export function checkCollision(boxA: any, boxB: any): boolean;
export function isContained(inner: any, outer: any): boolean;
export function alignShapes(shapeIdsOrBoxes: any[], alignment: string): any;
export function distributeShapes(shapeIdsOrBoxes: any[], direction: string): any;
export function getBounds(element: any): { x: number; y: number; width: number; height: number };
export function setBounds(element: any, bounds: Partial<{ x: number; y: number; width: number; height: number }>): void;
export function move(element: any, dx: number, dy: number): void;
export function resize(element: any, dw: number, dh: number): void;

export function generateRelationshipId(slideId: any): string;
export function parseRelationshipId(relId: string): any;
export function validateXml(xmlStr: string): any;
export function validateXML(xmlStr: string): any;
export function safeParseXml(xmlStr: string): any;
export function repairXML(xmlStr: string): string;
export function scanForEntities(xmlStr: string): any;
export function analyzeXmlFile(xmlStr: string): any;
export function reportXmlComplexity(xmlStr: string): any;
export function createLogger(name: string): any;
export function setGlobalLogLevel(level: string): void;
export function resetLogLevel(): void;

export class PPTXError extends Error {}
export class SlideNotFoundError extends PPTXError {}
export class ChartNotFoundError extends PPTXError {}
export class TableNotFoundError extends PPTXError {}
