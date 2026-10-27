export type ActionPreset =
  | 'next'
  | 'nextSlide'
  | 'previous'
  | 'prev'
  | 'previousSlide'
  | 'prevSlide'
  | 'first'
  | 'firstSlide'
  | 'last'
  | 'lastSlide'
  | 'end'
  | 'endShow'
  | string;

export interface HyperlinkOptions {
  tooltip?: string;
  action?: ActionPreset;
  targetSlide?: number;
  isInternal?: boolean;
  slide?: number;
  url?: string | number;
  color?: string | number;
  fontColor?: string | number;
  textColor?: string | number;
  underline?: boolean | string;
  underlineColor?: string | number;
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  fontFace?: string;
  fontFamily?: string;
  highlight?: string | boolean;
  style?: Partial<HyperlinkOptions>;
}

export interface TextHyperlinkOptions extends HyperlinkOptions {
  color?: string | number;
  fontColor?: string | number;
  textColor?: string | number;
  underline?: boolean | string;
  underlineColor?: string | number;
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  fontFace?: string;
  fontFamily?: string;
  highlight?: string | boolean;
  style?: Partial<TextHyperlinkOptions>;
}

export interface HyperlinkDescriptor {
  targetId: string;
  url?: string;
  isInternal: boolean;
  targetSlide?: number;
  action?: string;
  tooltip?: string;
  type: 'shape' | 'textRun' | 'image';
  shapeName?: string;
  shapeId?: string;
  text?: string;
  rId?: string;
}

export interface TableCellStyle {
  fontSize?: number;
  fontFace?: string;
  fontFamily?: string;
  color?: string;
  fontColor?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | string;
  alignment?: 'left' | 'center' | 'right' | 'justify' | string;
  align?: 'left' | 'center' | 'right' | 'justify' | string;
  verticalAlignment?: 'top' | 'middle' | 'center' | 'bottom' | string;
  verticalAlign?: 'top' | 'middle' | 'center' | 'bottom' | string;
  valign?: 'top' | 'middle' | 'center' | 'bottom' | string;
  fill?: string | false | null;
  bgColor?: string;
  backgroundColor?: string;
  margins?: number | { left?: number; right?: number; top?: number; bottom?: number };
  border?: any;
  borders?: any;
  [key: string]: any;
}

export interface TableCellShapeOptions {
  shape: {
    type?: 'roundRect' | 'circle' | 'ellipse' | 'rect' | 'square' | 'triangle' | 'star5' | 'rightArrow' | string;
    preset?: string;
    fill?: string;
    stroke?: string;
    strokeColor?: string;
    strokeWidth?: number;
    lineWidth?: number;
    width?: number;
    height?: number;
    radius?: number;
    border?: any;
    [key: string]: any;
  };
  text?: string | number | boolean;
  textPosition?: 'inside' | 'outside' | 'none';
  textStyle?: TableCellStyle;
  shapePlacement?: 'left' | 'right';
  style?: TableCellStyle;
  [key: string]: any;
}

export interface TableCellStyledValue {
  value?: string | number | boolean;
  text?: string | number | boolean;
  style?: TableCellStyle;
  [key: string]: any;
}

export type TableCellPrimitive = string | number | boolean | null | undefined;
export type TableCellData = TableCellPrimitive | TableCellStyledValue | TableCellShapeOptions;

export class PPTXTemplater {
  static load(source: string | Buffer | any, options?: any): Promise<PPTXTemplater>;
  static createBlank(options?: any): Promise<PPTXTemplater>;
  static extractPptx(pptxPath: string, outputPath: string, options?: any): Promise<void>;
  static buildPptx(sourceDir: string, outputPath: string, options?: any): Promise<void>;
  static preload(source: string | Buffer | any, options?: any): Promise<any>;
  static fromCache(source: string | Buffer | any): Promise<PPTXTemplater>;
  static clearCache(): void;
  static setLogLevel(level: string): void;

  slideCount: number;
  useSlide(...slideRefs: (number | string)[]): this;
  useAllSlides(): this;
  addSlide(options?: any, position?: number): this;
  duplicateSlide(sourceIndex: number, atPosition?: number): Promise<this>;
  removeSlide(position: number): this;
  getSlides(): any[];
  getSlide(index: number | string): any;
  importSlideFrom(sourcePpt: PPTXTemplater, sourceIndex: number | string, atPosition?: number): Promise<this>;
  exportSlides(...slideIndices: number[]): Promise<PPTXTemplater>;
  moveSlide(fromIndex: number, toIndex: number): this;
  reorderSlides(order: number[]): this;
  tagSlide(slideIndex: number, tag: string): this;
  getSlideDimensions(): { width: number; height: number };
  getSlideLayout(slideIndex?: number): string | null;
  useSection(sectionRef: string | number): this;
  addSection(name: string, startSlideIndex?: number): this;
  updateSection(sectionIdentifier: string | number, options: { name?: string }): this;
  removeSection(sectionIdentifier: string | number, removeSlides?: boolean): this;
  getSection(sectionRef: string | number): any;
  hideSlide(...slideRefs: (number | string)[]): this;
  unhideSlide(...slideRefs: (number | string)[]): this;
  resetSlide(slideRef?: number | string | { preservePlaceholders?: boolean }, options?: { preservePlaceholders?: boolean }): this;
  addComment(text: string, options?: { author?: string; initials?: string; x?: number | string; y?: number | string; dateTime?: string }): Promise<this>;
  removeComment(commentIdentifier: number | string, options?: { author?: string }): Promise<this>;
  checkCollision(idA: string, idB: string): boolean;
  isContained(innerId: string, outerId: string): boolean;
  replaceText(replacements: Record<string, string | number>): this;
  addText(textContent: string, options?: any): this;
  replaceTextByTag(tagOrId: string | number, newText: string | number, options?: any): this;
  updateText(tagOrId: string | number, newText: string | number | any, options?: any): this;
  addLists(items: any[], options?: any): this;
  getList(tag: string): any[];
  findText(query: string, options?: any): any[];
  addStyle(target: string, style: any): this;
  removeStyle(target: string, styleProperties: string[] | any): this;
  updateStyle(target: string, style: any): this;
  getText(tag: string): string;
  getLists(): any[];
  findLists(query: string): any[];
  removeText(tag: string): this;
  removeList(tag: string): this;
  clearText(tag: string): this;
  clearList(tag: string): this;
  replaceList(tag: string, newItems: any[] | any, options?: any): this;
  getTextElements(): any[];
  // Charts (Strict 11 Whitelisted Methods)
  addChart(chartType: string, data: any, options?: any): this;
  addChart(slideId: number | string, chartType: string, data: any, options?: any): this;
  updateChart(chartId: string, options: any): this;
  updateChartData(chartId: string, data: any): this;
  removeChart(chartId: string): this;
  clearCharts(slideRefOrChartId?: number | string): this;
  getCharts(slideId?: number | string): any[];
  getChart(chartId: string): any;
  cloneChart(chartId: string, clonedChartId?: string, clonedOptions?: any): this;
  // Hyperlinks (Strict 6 Whitelisted Methods)
  addHyperlink(targetId: string | number, urlOrAction: string | number | ActionPreset | HyperlinkOptions, options?: HyperlinkOptions): this;
  updateHyperlink(targetId: string | number, urlOrAction?: string | number | ActionPreset | HyperlinkOptions, options?: HyperlinkOptions): this;
  removeHyperlink(targetId: string | number | { targetId?: string | number; shapeName?: string; text?: string; slide?: number }): this;
  getHyperlinks(slideId?: number): HyperlinkDescriptor[];
  getHyperlink(targetId: string | number): HyperlinkDescriptor | null;
  setHyperlink(targetId: string | number | null, text: string, urlOrAction: string | number | ActionPreset | TextHyperlinkOptions, options?: TextHyperlinkOptions): this;
  // Tables (Strict 30 Whitelisted Methods)
  addTable(slideId: number | string, rows: number, columns: number, data?: TableCellData[][], options?: any): this;
  addTable(rows: number, columns: number, data?: TableCellData[][], options?: any): this;
  addTable(options: any): this;
  updateTable(tableId: string, options: any): this;
  updateTableData(tableId: string, data: TableCellData[][]): this;
  removeTable(tableId: string): this;
  getTables(slideId?: number | string): any[];
  getTable(tableId: string): any;
  cloneTable(tableId: string, clonedTableId: string, clonedOptions?: any): this;
  addTableRow(tableId: string, rowIndex?: number | TableCellData[], data?: TableCellData[]): this;
  addTableColumn(tableId: string, columnIndex?: number | TableCellData[], data?: TableCellData[]): this;
  removeTableRow(tableId: string, rowIndex: number): this;
  removeTableColumn(tableId: string, columnIndex: number): this;
  updateRowData(tableId: string, rowIndex: number, data: TableCellData[]): this;
  updateColumnData(tableId: string, columnIndex: number, data: TableCellData[]): this;
  updateTableCell(tableId: string, rowIndex: number, columnIndex: number, options: TableCellData | TableCellStyle): this;
  clearRowData(tableId: string, rowIndex: number): this;
  clearColumnData(tableId: string, columnIndex: number): this;
  clearCellData(tableId: string, rowIndex: number, columnIndex: number): this;
  getTableData(tableId: string): string[][];
  getColumnData(tableId: string, columnIndex: number): string[];
  getRowData(tableId: string, rowIndex: number): string[];
  getColumn(tableId: string, columnIndex: number): any;
  getRow(tableId: string, rowIndex: number): any;
  getCell(tableId: string, rowIndex: number, columnIndex: number): any;
  getCellData(tableId: string, rowIndex: number, columnIndex: number): string;
  mergeCells(tableId: string, startRow: number, startColumn: number, endRow: number, endColumn: number): this;
  splitCell(tableId: string, rowIndex: number, columnIndex: number): this;
  styleTable(tableId: string, options: any): this;
  styleColumn(tableId: string, columnIndex: number, options: any): this;
  styleRow(tableId: string, rowIndex: number, options: any): this;
  styleCell(tableId: string, rowIndex: number, columnIndex: number, options: any): this;
  saveToFile(filePath: string, options?: any): Promise<void>;
  toBuffer(options?: any): Promise<Buffer>;
  toStream(options?: any): Promise<any>;
  toBase64(options?: { dataUri?: boolean }): Promise<string>;
  toArrayBuffer(options?: any): Promise<ArrayBuffer>;
  toResponse(options?: any): Promise<Response>;
  toExpress(res: any, filename?: string, options?: any): Promise<void>;
  toFastify(reply: any, filename?: string, options?: any): Promise<void>;
  extractToZip(destinationPath: string, options?: any): Promise<void>;
  extractToXml(destinationDir: string, options?: any): Promise<void>;
  toFlatXml(destinationPathOrOptions?: string | any, options?: any): Promise<string>;
  repair(): Promise<this>;
  inspectXML(xmlPath: string): Promise<any>;
  getDataLabels(chartId: string, options?: any): Promise<any>;
  validateDataLabels(chartId: string, options?: any): Promise<any>;
  validateChartLabels(chartId: string, options?: any): Promise<any>;
  validateSeriesNameLabels(chartId: string, options?: any): Promise<any>;
  getChartLabelPositions(chartId: string): Promise<any>;
  addShape(slideId: number | string, type: string, options?: any): this;
  addShape(typeOrOptions: any, options?: any): this;
  updateShape(targetId: string, options: any): this;
  removeShape(shapeId: string): this;
  clearShapes(): this;
  getShapes(): any[];
  getShape(shapeId: string): any;
  cloneShape(shapeId: string, clonedShapeId: string, clonedOptions?: any): this;
  getGroupShapes(groupId: string): any[];
  groupShapes(shapeIds: string[], groupOptions?: any): this;
  ungroupShapes(groupId: string): this;
  addImage(slideId: number | string, source: string | Buffer | Uint8Array, options?: any): Promise<this>;
  addImage(source: string | Buffer | Uint8Array, options?: any): Promise<this>;
  updateImage(imageId: string | number, options?: any): Promise<this>;
  removeImage(imageId: string | number): this;
  addVideo(slideId: number | string, source: string | Buffer | Uint8Array, options?: any): Promise<this>;
  addVideo(source: string | Buffer | Uint8Array, options?: any): Promise<this>;
  updateVideo(videoId: string | number, options?: any): Promise<this>;
  removeVideo(videoId: string | number): this;
  addAudio(slideId: number | string, source: string | Buffer | Uint8Array, options?: any): Promise<this>;
  addAudio(source: string | Buffer | Uint8Array, options?: any): Promise<this>;
  updateAudio(audioId: string | number, options?: any): Promise<this>;
  removeAudio(audioId: string | number): this;
  getMedias(slideId?: number | string): any[];
  getMedia(mediaId: string | number): any;
  cloneMedia(mediaId: string | number, clonedMediaId?: string | number, clonedOptions?: any): this;
  AddImage(slideId: number | string, source: string | Buffer | Uint8Array, options?: any): Promise<this>;
  AddImage(source: string | Buffer | Uint8Array, options?: any): Promise<this>;
  UpdateImage(imageId: string | number, options?: any): Promise<this>;
  RemoveImage(imageId: string | number): this;
  AddVideo(slideId: number | string, source: string | Buffer | Uint8Array, options?: any): Promise<this>;
  AddVideo(source: string | Buffer | Uint8Array, options?: any): Promise<this>;
  UpdateVideo(videoId: string | number, options?: any): Promise<this>;
  RemoveVideo(videoId: string | number): this;
  AddAudio(slideId: number | string, source: string | Buffer | Uint8Array, options?: any): Promise<this>;
  AddAudio(source: string | Buffer | Uint8Array, options?: any): Promise<this>;
  UpdateAudio(audioId: string | number, options?: any): Promise<this>;
  RemoveAudio(audioId: string | number): this;
  GetMedias(slideId?: number | string): any[];
  GetMedia(mediaId: string | number): any;
  CloneMedia(mediaId: string | number, clonedMediaId?: string | number, clonedOptions?: any): this;
  // Elements (Strict 9 Whitelisted Methods)
  getElements(slideIdOrFilter?: number | string | ElementFilterOptions, filter?: ElementFilterOptions): ElementDescriptor[];
  getElement(elementId: string | number, slideId?: number | string): ElementDescriptor | null;
  removeElement(elementIdOrOptions: string | number | { slide?: number | string; elementId?: string | number; objectId?: string | number; id?: string | number }, slideId?: number | string): this;
  removeAllElements(slideIdOrOptions?: number | string | { slide?: number | string; preservePlaceholders?: boolean }, options?: { preservePlaceholders?: boolean }): this;
  setZOrder(optionsOrId: string | number | { slide?: number | string; elementId?: string | number; objectId?: string | number; id?: string | number; zIndex?: number }, zIndex?: number, slideId?: number | string): this;
  bringToFront(optionsOrId: string | number | { slide?: number | string; elementId?: string | number; objectId?: string | number; id?: string | number }, slideId?: number | string): this;
  sendToBack(optionsOrId: string | number | { slide?: number | string; elementId?: string | number; objectId?: string | number; id?: string | number }, slideId?: number | string): this;
  bringForward(optionsOrId: string | number | { slide?: number | string; elementId?: string | number; objectId?: string | number; id?: string | number }, slideId?: number | string): this;
  sendBackward(optionsOrId: string | number | { slide?: number | string; elementId?: string | number; objectId?: string | number; id?: string | number }, slideId?: number | string): this;
  validatePresentation(): Promise<any>;
  validatePresentationXml(): Promise<any>;
  validateSlide(slideIndex: number): Promise<any>;
  validateTable(tableId: string): Promise<any>;
  validateArchive(): Promise<any>;
}

export const PPTXTemplate: typeof PPTXTemplater;

export class PackageLoader {
  static load(source: any, options?: any): Promise<VirtualFileSystem>;
}
export function isPkZip(buf: Uint8Array | Buffer): boolean;
export function isLegacyPpt(buf: Uint8Array | Buffer): boolean;
export function isOdp(buf: Uint8Array | Buffer): Promise<boolean>;
export function isFlatOpc(source: any): boolean;
export class FlatOpcConverter {
  static convert(source: string | Buffer | Uint8Array, options?: any): VirtualFileSystem;
  static serialize(vfs: VirtualFileSystem, options?: any): string;
  static inferContentType(vfsPath: string, vfs?: VirtualFileSystem): string;
}
export class LegacyFormatConverter {
  static isLegacyPpt(buf: Uint8Array | Buffer): boolean;
  static isOdp(buf: Uint8Array | Buffer): Promise<boolean>;
  static supports(buf: Uint8Array | Buffer): Promise<boolean>;
  static detectFormat(buf: Uint8Array | Buffer): Promise<'ppt' | 'odp' | 'openxml' | 'unknown'>;
  static convert(buf: Uint8Array | Buffer, options?: any): Promise<VirtualFileSystem>;
}
export function formatXml(xml: string): string;
export function createBlankPresentationVfs(options?: any): VirtualFileSystem;
export function sendExpress(res: any, streamOrBuffer: any, filename?: string, options?: any): void;
export function sendFastify(reply: any, streamOrBuffer: any, filename?: string, options?: any): Promise<any>;
export function sendKoa(ctx: any, streamOrBuffer: any, filename?: string, options?: any): void;
export function createWebResponse(bufferOrArrayBuffer: any, filename?: string, options?: any): Response;

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

export class CellShapeManager {
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

export class ElementManager {
  [key: string]: any;
}

export interface ElementDescriptor {
  id: string;
  name: string;
  type: string;
  bounds: {
    x: number;
    y: number;
    width: number;
    height: number;
    rotation?: number;
    flipH?: boolean;
    flipV?: boolean;
  };
  slideIndex: number;
  element: any;
  frame?: any;
}

export interface ElementFilterOptions {
  type?: string;
  name?: string | RegExp;
  id?: string | number;
  slideIndex?: number;
  slide?: number | string;
  bounds?: any;
  within?: any;
  intersects?: any;
  overlaps?: any;
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
