/**
 * @fileoverview LegacyFormatConverter — binary PPT and ODP normalization engine.
 *
 * Provides header detection, stream inspection, and normalization of legacy
 * Microsoft PowerPoint (.ppt) and OpenDocument Presentation (.odp) files
 * into valid OpenXML PresentationML VirtualFileSystem packages.
 *
 * @module LegacyFormatConverter
 */

'use strict'

const JSZip = require('jszip')
const { createBlankPresentationVfs } = require('../templates/blankPresentation.js')
const { createLogger } = require('../utils/logger.js')

const logger = createLogger('LegacyFormatConverter')

/**
 * Microsoft Compound File Binary Format (CFBF / OLE2) magic signature:
 * 0xD0 0xCF 0x11 0xE0 0xA1 0xB1 0x1A 0xE1
 */
const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]

/**
 * PKZip local file header magic: PK\x03\x04
 */
const PKZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04]

/**
 * Checks if the buffer matches the OLE CFBF / legacy PPT magic header.
 *
 * @param {Buffer|Uint8Array} buf
 * @returns {boolean}
 */
function isLegacyPpt(buf) {
  if (!buf || buf.length < 8) return false
  for (let i = 0; i < 8; i++) {
    if (buf[i] !== OLE_MAGIC[i]) return false
  }
  return true
}

/**
 * Checks if the buffer matches the PKZip magic header.
 *
 * @param {Buffer|Uint8Array} buf
 * @returns {boolean}
 */
function isPkZip(buf) {
  if (!buf || buf.length < 4) return false
  return (
    buf[0] === PKZIP_MAGIC[0] &&
    buf[1] === PKZIP_MAGIC[1] &&
    buf[2] === PKZIP_MAGIC[2] &&
    buf[3] === PKZIP_MAGIC[3]
  )
}

/**
 * Checks if a buffer represents an OpenDocument Presentation (.odp).
 *
 * @param {Buffer|Uint8Array} buf
 * @returns {Promise<boolean>}
 */
async function isOdp(buf) {
  if (!isPkZip(buf)) return false
  try {
    const zip = await JSZip.loadAsync(buf)
    const mimetypeFile = zip.file('mimetype')
    if (mimetypeFile) {
      const mime = (await mimetypeFile.async('text')).trim()
      if (mime.includes('oasis.opendocument.presentation') || mime.includes('opendocument.presentation')) {
        return true
      }
    }
    const contentFile = zip.file('content.xml')
    if (contentFile) {
      const content = await contentFile.async('text')
      if (content.includes('urn:oasis:names:tc:opendocument:xmlns:presentation:1.0') ||
          content.includes('office:presentation') ||
          content.includes('draw:page')) {
        return true
      }
    }
    return false
  } catch {
    return false
  }
}

/**
 * Extracts printable text strings from a binary PPT buffer.
 *
 * @param {Buffer|Uint8Array} buf
 * @returns {string[]}
 */
function extractBinaryStrings(buf) {
  const buffer = Buffer.isBuffer(buf) ? buf : Buffer.from(buf)
  const strings = []

  // 1. Scan for ASCII sequences (min length 4)
  let currentAscii = ''
  for (let i = 0; i < buffer.length; i++) {
    const byte = buffer[i]
    if (byte >= 32 && byte <= 126) {
      currentAscii += String.fromCharCode(byte)
    } else {
      if (currentAscii.trim().length >= 4) {
        // Filter out binary garbage / keywords
        const trimmed = currentAscii.trim()
        if (!/^[a-zA-Z0-9_\-./\s,:;()]{4,}$/.test(trimmed)) {
          // ignore
        } else if (!strings.includes(trimmed)) {
          strings.push(trimmed)
        }
      }
      currentAscii = ''
    }
  }

  // 2. Scan for UTF-16LE sequences (min length 4)
  let currentWide = ''
  for (let i = 0; i < buffer.length - 1; i += 2) {
    const b1 = buffer[i]
    const b2 = buffer[i + 1]
    if (b2 === 0 && b1 >= 32 && b1 <= 126) {
      currentWide += String.fromCharCode(b1)
    } else {
      if (currentWide.trim().length >= 4) {
        const trimmed = currentWide.trim()
        if (!strings.includes(trimmed)) {
          strings.push(trimmed)
        }
      }
      currentWide = ''
    }
  }

  return strings.filter(s => {
    // Exclude common OLE internal stream names
    const lower = s.toLowerCase()
    return (
      !lower.startsWith('current user') &&
      !lower.startsWith('powerpoint document') &&
      !lower.startsWith('summaryinformation') &&
      !lower.startsWith('documentinformation')
    )
  })
}

/**
 * LegacyFormatConverter handles detection and normalization of .ppt and .odp formats.
 */
class LegacyFormatConverter {
  /**
   * Detects whether source buffer is a legacy PPT format.
   *
   * @param {Buffer|Uint8Array} buf
   * @returns {boolean}
   */
  static isLegacyPpt(buf) {
    return isLegacyPpt(buf)
  }

  /**
   * Detects whether source buffer is an OpenDocument Presentation (.odp).
   *
   * @param {Buffer|Uint8Array} buf
   * @returns {Promise<boolean>}
   */
  static async isOdp(buf) {
    return isOdp(buf)
  }

  /**
   * Returns true if the buffer is a supported legacy/alternative format (.ppt or .odp).
   *
   * @param {Buffer|Uint8Array} buf
   * @returns {Promise<boolean>}
   */
  static async supports(buf) {
    if (!buf || buf.length < 4) return false
    if (isLegacyPpt(buf)) return true
    return isOdp(buf)
  }

  /**
   * Detects format name from buffer bytes.
   *
   * @param {Buffer|Uint8Array} buf
   * @returns {Promise<'ppt'|'odp'|'openxml'|'unknown'>}
   */
  static async detectFormat(buf) {
    if (!buf || buf.length < 4) return 'unknown'
    if (isLegacyPpt(buf)) return 'ppt'
    if (isPkZip(buf)) {
      if (await isOdp(buf)) return 'odp'
      return 'openxml'
    }
    return 'unknown'
  }

  /**
   * Converts a legacy PPT or ODP presentation into an OpenXML VirtualFileSystem.
   *
   * @param {Buffer|Uint8Array} buf
   * @param {Object} [options={}]
   * @returns {Promise<VirtualFileSystem>}
   */
  static async convert(buf, options = {}) {
    if (isLegacyPpt(buf)) {
      logger.info('Converting legacy binary PPT format to OpenXML package...')
      return LegacyFormatConverter.#convertPptToOpenXml(buf, options)
    }

    if (await isOdp(buf)) {
      logger.info('Converting ODP format to OpenXML package...')
      return LegacyFormatConverter.#convertOdpToOpenXml(buf, options)
    }

    throw new Error('LegacyFormatConverter: Unsupported legacy format or signature.')
  }

  /**
   * Converts legacy binary PPT into an OpenXML VirtualFileSystem.
   *
   * @private
   * @param {Buffer|Uint8Array} buf
   * @param {Object} options
   * @returns {VirtualFileSystem}
   */
  static #convertPptToOpenXml(buf, options = {}) {
    const vfs = createBlankPresentationVfs(options)
    const extractedStrings = extractBinaryStrings(buf)

    // Build a slide containing extracted texts if any
    const title = extractedStrings[0] || 'Converted Legacy Presentation'
    const bullets = extractedStrings.slice(1, 15) // take top extracted text items

    const paragraphs = bullets
      .map(
        b => `<a:p>
          <a:pPr lvl="0"/>
          <a:r>
            <a:rPr lang="en-US" sz="1800"/>
            <a:t>${escapeXml(b)}</a:t>
          </a:r>
        </a:p>`
      )
      .join('\n')

    const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
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
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="838200" y="365125"/>
            <a:ext cx="10515600" cy="1325563"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="4000" b="1"/>
              <a:t>${escapeXml(title)}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Content Placeholder 2"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="838200" y="1825625"/>
            <a:ext cx="10515600" cy="4351338"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          ${paragraphs || '<a:p><a:r><a:t>(Converted from legacy PPT)</a:t></a:r></a:p>'}
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`

    vfs.writeText('ppt/slides/slide1.xml', slideXml)
    return vfs
  }

  /**
   * Converts ODP ZIP archive into an OpenXML VirtualFileSystem.
   *
   * @private
   * @param {Buffer|Uint8Array} buf
   * @param {Object} options
   * @returns {Promise<VirtualFileSystem>}
   */
  static async #convertOdpToOpenXml(buf, options = {}) {
    const vfs = createBlankPresentationVfs(options)
    const zip = await JSZip.loadAsync(buf)

    const contentXmlFile = zip.file('content.xml')
    if (!contentXmlFile) {
      return vfs
    }

    const contentXml = await contentXmlFile.async('text')
    // Extract pages / slides from content.xml
    const pages = extractOdpPages(contentXml)

    if (pages.length === 0) {
      return vfs
    }

    // Update slide1 with first page content
    const firstPage = pages[0]
    vfs.writeText('ppt/slides/slide1.xml', createSlideFromOdpPage(firstPage, 1))

    // If multiple pages, add additional slides
    if (pages.length > 1) {
      const presentationXml = vfs.readText('ppt/presentation.xml')
      const presentationRels = vfs.readText('ppt/_rels/presentation.xml.rels')
      const contentTypes = vfs.readText('[Content_Types].xml')

      let sldIdEntries = '    <p:sldId id="256" r:id="rId2"/>\n'
      let relEntries = '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>\n  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>\n  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>\n'
      let contentTypesOverrides = ''

      for (let i = 1; i < pages.length; i++) {
        const slideIndex = i + 1
        const rId = `rId${slideIndex + 2}`
        const sldId = 256 + i
        const slidePath = `ppt/slides/slide${slideIndex}.xml`
        const slideRelsPath = `ppt/slides/_rels/slide${slideIndex}.xml.rels`

        vfs.writeText(slidePath, createSlideFromOdpPage(pages[i], slideIndex))
        vfs.writeText(
          slideRelsPath,
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`
        )

        sldIdEntries += `    <p:sldId id="${sldId}" r:id="${rId}"/>\n`
        relEntries += `  <Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideIndex}.xml"/>\n`
        contentTypesOverrides += `  <Override PartName="/${slidePath}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>\n`
      }

      // Update presentation.xml
      const updatedPres = presentationXml.replace(
        /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
        `<p:sldIdLst>\n${sldIdEntries}  </p:sldIdLst>`
      )
      vfs.writeText('ppt/presentation.xml', updatedPres)

      // Update presentation.xml.rels
      const updatedPresRels = presentationRels.replace(
        /<Relationships[\s\S]*?<\/Relationships>/,
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n${relEntries}</Relationships>`
      )
      vfs.writeText('ppt/_rels/presentation.xml.rels', updatedPresRels)

      // Update [Content_Types].xml
      const updatedContentTypes = contentTypes.replace(
        '</Types>',
        `${contentTypesOverrides}</Types>`
      )
      vfs.writeText('[Content_Types].xml', updatedContentTypes)
    }

    return vfs
  }
}

/**
 * Parses <draw:page> elements from ODP content.xml.
 *
 * @param {string} contentXml
 * @returns {Array<{ title: string, texts: string[] }>}
 */
function extractOdpPages(contentXml) {
  const pages = []
  const pageRegex = /<draw:page\b[^>]*>([\s\S]*?)<\/draw:page>/g
  let match

  while ((match = pageRegex.exec(contentXml)) !== null) {
    const pageContent = match[1]
    const pRegex = /<text:p\b[^>]*>([\s\S]*?)<\/text:p>/g
    const texts = []
    let pMatch

    while ((pMatch = pRegex.exec(pageContent)) !== null) {
      const rawText = pMatch[1].replace(/<[^>]+>/g, '').trim()
      if (rawText) {
        texts.push(rawText)
      }
    }

    const title = texts[0] || 'Slide'
    const bodyTexts = texts.slice(1)
    pages.push({ title, texts: bodyTexts })
  }

  return pages
}

/**
 * Builds OpenXML slide XML from an ODP page object.
 *
 * @param {{ title: string, texts: string[] }} page
 * @param {number} slideIndex
 * @returns {string}
 */
function createSlideFromOdpPage(page, slideIndex) {
  const title = page.title || `Slide ${slideIndex}`
  const paragraphs = page.texts
    .map(
      t => `<a:p>
        <a:pPr lvl="0"/>
        <a:r>
          <a:rPr lang="en-US" sz="1800"/>
          <a:t>${escapeXml(t)}</a:t>
        </a:r>
      </a:p>`
    )
    .join('\n')

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
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
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="838200" y="365125"/>
            <a:ext cx="10515600" cy="1325563"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="4000" b="1"/>
              <a:t>${escapeXml(title)}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Content Placeholder 2"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="838200" y="1825625"/>
            <a:ext cx="10515600" cy="4351338"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          ${paragraphs || '<a:p><a:r><a:t>(Converted from ODP)</a:t></a:r></a:p>'}
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`
}

function escapeXml(str) {
  if (!str) return ''
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

module.exports = {
  LegacyFormatConverter,
  isLegacyPpt,
  isOdp,
  isPkZip,
  OLE_MAGIC,
  PKZIP_MAGIC,
}
