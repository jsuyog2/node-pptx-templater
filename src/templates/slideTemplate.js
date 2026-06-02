/**
 * @fileoverview Slide XML template builder.
 *
 * Generates the XML for new slides added via addSlide().
 * Produces minimal but valid OpenXML slide structures.
 *
 * OpenXML Slide XML Structure:
 * ─────────────────────────────────────────────────────────────────
 * <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
 * <p:sld xmlns:a="...drawingml..." xmlns:r="...relationships..." xmlns:p="...presentationml...">
 *   <p:cSld>                         ← Common Slide Data
 *     <p:spTree>                     ← Shape Tree (contains all shapes)
 *       <p:nvGrpSpPr>...</p:nvGrpSpPr>  ← Group shape properties
 *       <p:grpSpPr>...</p:grpSpPr>      ← Group visual properties
 *       <p:sp>...</p:sp>                ← Individual shapes
 *     </p:spTree>
 *   </p:cSld>
 *   <p:clrMapOvr>...</p:clrMapOvr>  ← Color map override (uses master)
 * </p:sld>
 *
 * Shape (p:sp) structure:
 *   <p:sp>
 *     <p:nvSpPr>                 ← Non-visual shape properties
 *       <p:cNvPr id="2" name="Title 1"/>
 *       <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
 *       <p:nvPr><p:ph type="title"/></p:nvPr>  ← Placeholder type
 *     </p:nvSpPr>
 *     <p:spPr/>                  ← Visual properties (inherit from layout)
 *     <p:txBody>                 ← Text body
 *       <a:bodyPr/>
 *       <a:lstStyle/>
 *       <a:p>                    ← Paragraph
 *         <a:r>                  ← Text run
 *           <a:t>Hello</a:t>     ← Text content
 *         </a:r>
 *       </a:p>
 *     </p:txBody>
 *   </p:sp>
 */

/**
 * OpenXML namespace declarations used in slide XML.
 */
const SLIDE_NAMESPACES = [
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"',
].join(' ')

/**
 * EMU conversion constants.
 */
const EMU = {
  /** 1 inch in EMU */
  INCH: 914400,
  /** Slide width (10 inches = 9144000 EMU) */
  SLIDE_WIDTH: 9144000,
  /** Slide height (7.5 inches = 6858000 EMU for 4:3, 5143500 for 16:9) */
  SLIDE_HEIGHT: 5143500,
}

/**
 * Builds the XML for a new slide.
 *
 * @param {NewSlideOptions} options - Slide configuration.
 * @param {number} slideIndex - 1-based slide index (for unique IDs).
 * @returns {string} Complete slide XML string.
 */
function buildNewSlideXml(options, _slideIndex) {
  const { title = '', elements = [], _layout = 'blank', _rawXml } = options

  // If raw XML is provided, use it directly (for clone/export operations)
  if (_rawXml) return _rawXml

  const shapes = []
  let shapeIdCounter = 1

  // Add title shape if title is provided
  if (title) {
    shapes.push(buildTitleShape(title, shapeIdCounter++))
  }

  // Add element shapes
  for (const element of elements) {
    shapes.push(buildElementShape(element, shapeIdCounter++))
  }

  const spTreeContent = [buildGroupShapeProps(), ...shapes].join('\n    ')

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${SLIDE_NAMESPACES} show="1">
  <p:cSld>
    <p:spTree>
      ${spTreeContent}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`
}

/**
 * Builds the group shape properties (required root element of spTree).
 * @returns {string}
 */
function buildGroupShapeProps() {
  return `<p:nvGrpSpPr>
      <p:cNvPr id="1" name=""/>
      <p:cNvGrpSpPr/>
      <p:nvPr/>
    </p:nvGrpSpPr>
    <p:grpSpPr>
      <a:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
        <a:chOff x="0" y="0"/>
        <a:chExt cx="0" cy="0"/>
      </a:xfrm>
    </p:grpSpPr>`
}

/**
 * Builds a title text shape.
 *
 * @param {string} title - Title text.
 * @param {number} shapeId - Unique shape ID.
 * @returns {string} XML for the title shape.
 */
function buildTitleShape(title, shapeId) {
  return `<p:sp>
      <p:nvSpPr>
        <p:cNvPr id="${shapeId}" name="Title ${shapeId}"/>
        <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
        <p:nvPr><p:ph type="title"/></p:nvPr>
      </p:nvSpPr>
      <p:spPr/>
      <p:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:r>
            <a:rPr lang="en-US" dirty="0"/>
            <a:t>${escapeXml(title)}</a:t>
          </a:r>
        </a:p>
      </p:txBody>
    </p:sp>`
}

/**
 * Builds a shape element based on its type.
 *
 * @param {SlideElement} element - Element descriptor.
 * @param {number} shapeId - Unique shape ID.
 * @returns {string} XML for the element.
 */
function buildElementShape(element, shapeId) {
  switch (element.type) {
    case 'text':
      return buildTextShape(element, shapeId)
    case 'image':
      return buildImagePlaceholder(element, shapeId)
    case 'shape':
      return buildBasicShape(element, shapeId)
    default:
      return buildTextShape({ ...element, value: element.value || '' }, shapeId)
  }
}

/**
 * Builds a text box shape.
 *
 * @param {TextElement} element
 * @param {number} shapeId
 * @returns {string}
 */
function buildTextShape(element, shapeId) {
  const {
    value = '',
    x = EMU.INCH,
    y = EMU.INCH,
    width = 6 * EMU.INCH,
    height = 1 * EMU.INCH,
    fontSize = 1800, // 18pt in hundredths of a point
    bold = false,
    italic = false,
    name = `TextBox ${shapeId}`,
  } = element

  const boldAttr = bold ? ' b="1"' : ''
  const italicAttr = italic ? ' i="1"' : ''

  return `<p:sp>
      <p:nvSpPr>
        <p:cNvPr id="${shapeId}" name="${escapeXml(name)}"/>
        <p:cNvSpPr txBox="1"/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="${x}" y="${y}"/>
          <a:ext cx="${width}" cy="${height}"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap="square" rtlCol="0">
          <a:normAutofit/>
        </a:bodyPr>
        <a:lstStyle/>
        <a:p>
          <a:r>
            <a:rPr lang="en-US" sz="${fontSize}" dirty="0"${boldAttr}${italicAttr}/>
            <a:t>${escapeXml(String(value))}</a:t>
          </a:r>
        </a:p>
      </p:txBody>
    </p:sp>`
}

/**
 * Builds an image placeholder shape.
 * Note: actual image embedding requires RelationshipManager.
 *
 * @param {ImageElement} element
 * @param {number} shapeId
 * @returns {string}
 */
function buildImagePlaceholder(element, shapeId) {
  const {
    rId = 'rId99', // Placeholder — caller must update this
    x = EMU.INCH,
    y = EMU.INCH,
    width = 3 * EMU.INCH,
    height = 2 * EMU.INCH,
    name = `Image ${shapeId}`,
  } = element

  return `<p:pic>
      <p:nvPicPr>
        <p:cNvPr id="${shapeId}" name="${escapeXml(name)}"/>
        <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
        <p:nvPr/>
      </p:nvPicPr>
      <p:blipFill>
        <a:blip r:embed="${rId}" cstate="print"/>
        <a:stretch><a:fillRect/></a:stretch>
      </p:blipFill>
      <p:spPr>
        <a:xfrm>
          <a:off x="${x}" y="${y}"/>
          <a:ext cx="${width}" cy="${height}"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
    </p:pic>`
}

/**
 * Builds a basic geometric shape.
 *
 * @param {ShapeElement} element
 * @param {number} shapeId
 * @returns {string}
 */
function buildBasicShape(element, shapeId) {
  const {
    shape = 'rect',
    x = EMU.INCH,
    y = EMU.INCH,
    width = 2 * EMU.INCH,
    height = EMU.INCH,
    fillColor = 'FFFFFF',
    name = `Shape ${shapeId}`,
  } = element

  return `<p:sp>
      <p:nvSpPr>
        <p:cNvPr id="${shapeId}" name="${escapeXml(name)}"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="${x}" y="${y}"/>
          <a:ext cx="${width}" cy="${height}"/>
        </a:xfrm>
        <a:prstGeom prst="${shape}"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="${fillColor}"/></a:solidFill>
      </p:spPr>
    </p:sp>`
}

/**
 * Escapes XML special characters.
 * @param {string} str
 * @returns {string}
 */
function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

module.exports = {
  buildNewSlideXml,
  EMU,
}
