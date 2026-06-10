/**
 * @fileoverview imageMetadata - Pure JS helper to read image dimensions without full file reads.
 */

const fs = require('fs')
const { promisify } = require('util')
const openAsync = promisify(fs.open)
const readAsync = promisify(fs.read)
const closeAsync = promisify(fs.close)

/**
 * Reads image dimensions and aspect ratio from a file path or buffer.
 *
 * @param {string|Buffer} source - File path or image Buffer.
 * @returns {Promise<{ width: number, height: number, aspectRatio: number, type: string }>}
 */
async function getImageMetadata(source) {
  let buffer
  const isBuffer = Buffer.isBuffer(source) || source instanceof Uint8Array

  if (isBuffer) {
    buffer = Buffer.isBuffer(source) ? source : Buffer.from(source)
  } else if (typeof source === 'string') {
    // Read only the first 8KB of the file
    let fd
    try {
      fd = await openAsync(source, 'r')
      const tempBuffer = Buffer.alloc(8192)
      const { bytesRead } = await readAsync(fd, tempBuffer, 0, 8192, 0)
      buffer = tempBuffer.subarray(0, bytesRead)
    } finally {
      if (fd !== undefined) {
        await closeAsync(fd).catch(() => {})
      }
    }
  } else {
    throw new Error('Unsupported image source type')
  }

  if (buffer.length < 4) {
    throw new Error('Image file is too small or corrupt')
  }

  // Detect image type by magic bytes
  if (buffer[0] === 0x89 && buffer[1] === 0x50 && buffer[2] === 0x4e && buffer[3] === 0x47) {
    return parsePng(buffer)
  }

  if (buffer[0] === 0xff && buffer[1] === 0xd8) {
    return parseJpeg(buffer)
  }

  if (buffer[0] === 0x47 && buffer[1] === 0x49 && buffer[2] === 0x46) {
    return parseGif(buffer)
  }

  if (buffer[0] === 0x42 && buffer[1] === 0x4d) {
    return parseBmp(buffer)
  }

  // Check for SVG (starts with XML declaration or <svg)
  const textContent = buffer.toString('utf8').trim()
  if (
    textContent.startsWith('<svg') ||
    textContent.includes('<svg') ||
    textContent.startsWith('<?xml')
  ) {
    return parseSvg(textContent)
  }

  throw new Error('Unsupported image format or unrecognized signature')
}

function parsePng(buffer) {
  if (buffer.length < 24) {
    throw new Error('PNG header too short')
  }
  // Width is at offset 16 (4 bytes, big endian)
  const width = buffer.readUInt32BE(16)
  // Height is at offset 20 (4 bytes, big endian)
  const height = buffer.readUInt32BE(20)

  return {
    width,
    height,
    aspectRatio: width / height,
    type: 'png',
  }
}

function parseGif(buffer) {
  if (buffer.length < 10) {
    throw new Error('GIF header too short')
  }
  // Width is at offset 6 (2 bytes, little endian)
  const width = buffer.readUInt16LE(6)
  // Height is at offset 8 (2 bytes, little endian)
  const height = buffer.readUInt16LE(8)

  return {
    width,
    height,
    aspectRatio: width / height,
    type: 'gif',
  }
}

function parseBmp(buffer) {
  if (buffer.length < 26) {
    throw new Error('BMP header too short')
  }
  // Width is at offset 18 (4 bytes, little endian)
  const width = buffer.readInt32LE(18)
  // Height is at offset 22 (4 bytes, little endian)
  const height = buffer.readInt32LE(22)

  return {
    width: Math.abs(width),
    height: Math.abs(height),
    aspectRatio: Math.abs(width) / Math.abs(height),
    type: 'bmp',
  }
}

function parseJpeg(buffer) {
  let offset = 2 // Skip SOI marker (FF D8)

  while (offset < buffer.length - 8) {
    // Check marker signature
    if (buffer[offset] !== 0xff) {
      // Not a valid marker, search next FF
      offset++
      continue
    }

    // Skip extra FF padding
    while (buffer[offset] === 0xff && offset < buffer.length) {
      offset++
    }

    if (offset >= buffer.length) break

    const marker = buffer[offset]
    offset++

    // SOI, EOI, TEM have no length
    if (marker === 0xd8 || marker === 0xd9 || marker === 0x01) {
      continue
    }

    // Read segment length (2 bytes, big endian)
    const length = buffer.readUInt16BE(offset)

    // Check SOF markers: C0-C3, C5-CB, CD-CF
    const isSOF =
      (marker >= 0xc0 && marker <= 0xc3) ||
      (marker >= 0xc5 && marker <= 0xcb) ||
      (marker >= 0xcd && marker <= 0xcf)

    if (isSOF) {
      // SOF structure:
      // Offset 0: precision (1 byte)
      // Offset 1: height (2 bytes, big endian)
      // Offset 3: width (2 bytes, big endian)
      const height = buffer.readUInt16BE(offset + 3)
      const width = buffer.readUInt16BE(offset + 5)

      return {
        width,
        height,
        aspectRatio: width / height,
        type: 'jpg',
      }
    }

    // Advance to next marker
    offset += length
  }

  throw new Error('Could not find JPEG SOF marker')
}

function parseSvg(text) {
  // Try finding <svg ...> tag
  const svgMatch = /<svg([^>]+)>/i.exec(text)
  if (!svgMatch) {
    throw new Error('Invalid SVG: missing <svg> tag')
  }

  const svgAttr = svgMatch[1]

  const widthMatch = /width\s*=\s*["']([^"']+)["']/i.exec(svgAttr)
  const heightMatch = /height\s*=\s*["']([^"']+)["']/i.exec(svgAttr)
  const viewBoxMatch = /viewBox\s*=\s*["']([^"']+)["']/i.exec(svgAttr)

  let width = 0
  let height = 0

  if (widthMatch) width = parseFloat(widthMatch[1])
  if (heightMatch) height = parseFloat(heightMatch[1])

  // If width/height missing or using units, fallback to viewBox
  if ((!width || !height || isNaN(width) || isNaN(height)) && viewBoxMatch) {
    const parts = viewBoxMatch[1].trim().split(/\s+/)
    if (parts.length === 4) {
      const vbWidth = parseFloat(parts[2])
      const vbHeight = parseFloat(parts[3])
      if (!isNaN(vbWidth) && !isNaN(vbHeight)) {
        width = width || vbWidth
        height = height || vbHeight
      }
    }
  }

  // Default fallbacks if everything fails
  width = width || 800
  height = height || 600

  return {
    width,
    height,
    aspectRatio: width / height,
    type: 'svg',
  }
}

module.exports = { getImageMetadata }
