/**
 * @fileoverview FlatOpcConverter — Ingestion and serialization engine for Microsoft Flat OPC XML packages.
 *
 * Implements ISO/IEC 29500-2 / ECMA-376 Part 2 Annex A Flat OPC (Open Packaging Conventions)
 * representation: single-file XML packages (<pkg:package>) containing XML parts (<pkg:xmlData>)
 * and Base64-encoded binary parts (<pkg:binaryData>).
 *
 * @module FlatOpcConverter
 */

'use strict'

const { VirtualFileSystem } = require('../core/VirtualFileSystem.js')
const { createLogger } = require('../utils/logger.js')
const { PPTXError } = require('../utils/errors.js')

const logger = createLogger('FlatOpcConverter')

const FLAT_OPC_NAMESPACE = 'http://schemas.microsoft.com/office/2006/xmlPackage'

/**
 * Standard MIME types by extension for synthesizing [Content_Types].xml
 */
const DEFAULT_CONTENT_TYPES = {
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
  xml: 'application/xml',
  png: 'image/png',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  bin: 'application/vnd.openxmlformats-officedocument.oleObject',
}

/**
 * Detects whether a string or Buffer contains Microsoft Flat OPC XML (<pkg:package>).
 *
 * @param {string|Buffer|Uint8Array} source
 * @returns {boolean}
 */
function isFlatOpc(source) {
  if (!source) return false

  let textHead = ''
  if (typeof source === 'string') {
    textHead = source.slice(0, 4096)
  } else if (Buffer.isBuffer(source) || source instanceof Uint8Array) {
    const sliceLen = Math.min(source.length, 4096)
    const subarray = Buffer.isBuffer(source)
      ? source.subarray(0, sliceLen)
      : Buffer.from(source.buffer, source.byteOffset, sliceLen)
    textHead = subarray.toString('utf8')
  } else {
    return false
  }

  return (
    textHead.includes('<pkg:package') ||
    textHead.includes(FLAT_OPC_NAMESPACE) ||
    /<\?mso-application\s+progid=["']PowerPoint\.(Show|Presentation)["']/i.test(textHead)
  )
}

/**
 * Converter for Microsoft Flat OPC XML packages.
 */
class FlatOpcConverter {
  /**
   * Converts a Flat OPC XML string or Buffer into a VirtualFileSystem.
   *
   * @param {string|Buffer|Uint8Array} source - Flat OPC XML content.
   * @param {Object} [options={}] - Conversion options.
   * @returns {VirtualFileSystem} Populated VirtualFileSystem instance.
   * @throws {PPTXError} If the XML structure is invalid or missing critical parts.
   */
  static convert(source, _options = {}) {
    logger.debug('Converting Flat OPC XML package into VirtualFileSystem')

    let xmlString
    if (typeof source === 'string') {
      xmlString = source
    } else if (Buffer.isBuffer(source)) {
      xmlString = source.toString('utf8')
    } else if (source instanceof Uint8Array) {
      xmlString = Buffer.from(source.buffer, source.byteOffset, source.byteLength).toString('utf8')
    } else {
      throw new PPTXError('FlatOpcConverter: source must be a string, Buffer, or Uint8Array')
    }

    const vfs = new VirtualFileSystem()
    const contentTypesOverrides = []

    // Match each <pkg:part ...>...</pkg:part> element
    const partRegex = /<pkg:part\s+([^>]+)>([\s\S]*?)<\/pkg:part>/gi
    let partMatch
    let partCount = 0

    while ((partMatch = partRegex.exec(xmlString)) !== null) {
      partCount++
      const rawAttrs = partMatch[1]
      const partBody = partMatch[2]

      const nameMatch = rawAttrs.match(/pkg:name=["']([^"']+)["']/i)
      const typeMatch = rawAttrs.match(/pkg:contentType=["']([^"']+)["']/i)

      if (!nameMatch) {
        logger.warn(`Skipping <pkg:part> without pkg:name attribute`)
        continue
      }

      const rawPartName = nameMatch[1]
      const contentType = typeMatch ? typeMatch[1] : ''

      // Strip leading slash for VFS storage: e.g. "/ppt/presentation.xml" -> "ppt/presentation.xml"
      const vfsPath = rawPartName.startsWith('/') ? rawPartName.slice(1) : rawPartName

      // 1. Binary part check (<pkg:binaryData>)
      const binMatch = partBody.match(/<pkg:binaryData[^>]*>([\s\S]*?)<\/pkg:binaryData>/i)
      if (binMatch) {
        const base64Content = binMatch[1].replace(/\s+/g, '')
        const binaryBuffer = Buffer.from(base64Content, 'base64')
        vfs.writeBytes(vfsPath, binaryBuffer)
        logger.debug(`Extracted binary part: ${vfsPath} (${binaryBuffer.length} bytes)`)
      } else {
        // 2. XML part check (<pkg:xmlData>)
        const xmlMatch = partBody.match(/<pkg:xmlData[^>]*>([\s\S]*?)<\/pkg:xmlData>/i)
        if (xmlMatch) {
          let innerXml = xmlMatch[1].trim()
          if (!innerXml.startsWith('<?xml')) {
            innerXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + innerXml
          }
          vfs.writeText(vfsPath, innerXml)
          logger.debug(`Extracted XML part: ${vfsPath} (${innerXml.length} chars)`)
        } else {
          // Empty or unrecognized part body
          vfs.writeText(vfsPath, '')
          logger.debug(`Extracted empty part: ${vfsPath}`)
        }
      }

      // Collect overrides for [Content_Types].xml (excluding relationships parts)
      if (
        contentType &&
        !vfsPath.endsWith('.rels') &&
        !vfsPath.endsWith('/_rels') &&
        vfsPath !== '[Content_Types].xml'
      ) {
        contentTypesOverrides.push({
          partName: `/${vfsPath}`,
          contentType,
        })
      }
    }

    if (partCount === 0) {
      throw new PPTXError('FlatOpcConverter: No <pkg:part> elements found in Flat OPC XML package')
    }

    // Synthesize [Content_Types].xml if the flat package did not include it as a separate part
    if (!vfs.has('[Content_Types].xml')) {
      logger.debug('Synthesizing [Content_Types].xml from Flat OPC part manifests')
      const defaultsXml = Object.entries(DEFAULT_CONTENT_TYPES)
        .map(([ext, type]) => `  <Default Extension="${ext}" ContentType="${type}"/>`)
        .join('\n')

      const overridesXml = contentTypesOverrides
        .map(o => `  <Override PartName="${o.partName}" ContentType="${o.contentType}"/>`)
        .join('\n')

      const contentTypesDoc =
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n` +
        `${defaultsXml}\n` +
        `${overridesXml}\n` +
        `</Types>`

      vfs.writeText('[Content_Types].xml', contentTypesDoc)
    }

    // Validate critical PresentationML parts
    if (!vfs.has('ppt/presentation.xml')) {
      throw new PPTXError(
        'FlatOpcConverter: Invalid Flat OPC package — missing critical part "ppt/presentation.xml"'
      )
    }

    logger.info(`Successfully parsed Flat OPC XML into VFS (${vfs.size} parts extracted)`)
    return vfs
  }

  /**
   * Serializes a VirtualFileSystem representation into a Microsoft Flat OPC XML package string.
   *
   * @param {VirtualFileSystem} vfs - In-memory presentation VFS.
   * @param {Object} [options={}] - Serialization options.
   * @returns {string} Complete Flat OPC XML string.
   */
  static serialize(vfs, _options = {}) {
    logger.debug('Serializing VirtualFileSystem into Flat OPC XML package')

    const lines = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<?mso-application progid="PowerPoint.Show"?>',
      `<pkg:package xmlns:pkg="${FLAT_OPC_NAMESPACE}">`,
    ]

    for (const [vfsPath, data] of vfs) {
      // Flat OPC packages don't typically include [Content_Types].xml as a part
      // because each part explicitly declares its own pkg:contentType.
      if (vfsPath === '[Content_Types].xml') continue

      const pkgName = vfsPath.startsWith('/') ? vfsPath : `/${vfsPath}`
      const contentType = FlatOpcConverter.inferContentType(vfsPath, vfs)
      const isBinary = data instanceof Uint8Array || Buffer.isBuffer(data)

      lines.push(`  <pkg:part pkg:name="${pkgName}" pkg:contentType="${contentType}">`)

      if (isBinary && !FlatOpcConverter.#isTextPart(vfsPath)) {
        const buf = Buffer.isBuffer(data)
          ? data
          : Buffer.from(data.buffer, data.byteOffset, data.byteLength)
        const base64 = buf.toString('base64')
        lines.push(`    <pkg:binaryData>${base64}</pkg:binaryData>`)
      } else {
        let textContent =
          typeof data === 'string'
            ? data
            : Buffer.from(data.buffer, data.byteOffset, data.byteLength).toString('utf8')

        // Strip leading <?xml ... ?> declaration inside <pkg:xmlData>
        textContent = textContent.replace(/^\s*<\?xml[^>]*\?>\s*/i, '').trim()
        lines.push(`    <pkg:xmlData>${textContent}</pkg:xmlData>`)
      }

      lines.push('  </pkg:part>')
    }

    lines.push('</pkg:package>')
    return lines.join('\n')
  }

  /**
   * Infers the OpenXML MIME type for a given VFS part path.
   *
   * @param {string} vfsPath
   * @param {VirtualFileSystem} [_vfs]
   * @returns {string}
   */
  static inferContentType(vfsPath, _vfs) {
    if (vfsPath.endsWith('.rels')) {
      return DEFAULT_CONTENT_TYPES.rels
    }
    if (vfsPath === 'ppt/presentation.xml') {
      return 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
    }
    if (vfsPath.startsWith('ppt/slides/slide') && vfsPath.endsWith('.xml')) {
      return 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
    }
    if (vfsPath.startsWith('ppt/slideLayouts/') && vfsPath.endsWith('.xml')) {
      return 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
    }
    if (vfsPath.startsWith('ppt/slideMasters/') && vfsPath.endsWith('.xml')) {
      return 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
    }
    if (vfsPath.startsWith('ppt/theme/') && vfsPath.endsWith('.xml')) {
      return 'application/vnd.openxmlformats-officedocument.theme+xml'
    }
    if (vfsPath.startsWith('ppt/charts/chart') && vfsPath.endsWith('.xml')) {
      return 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
    }
    if (vfsPath === 'docProps/core.xml') {
      return 'application/vnd.openxmlformats-package.core-properties+xml'
    }
    if (vfsPath === 'docProps/app.xml') {
      return 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
    }

    const ext = vfsPath.split('.').pop()?.toLowerCase()
    return DEFAULT_CONTENT_TYPES[ext] || 'application/xml'
  }

  static #isTextPart(vfsPath) {
    return (
      vfsPath.endsWith('.xml') ||
      vfsPath.endsWith('.rels') ||
      vfsPath.endsWith('.txt') ||
      vfsPath.endsWith('.json')
    )
  }
}

module.exports = {
  FlatOpcConverter,
  isFlatOpc,
  FLAT_OPC_NAMESPACE,
}
