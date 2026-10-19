/**
 * @fileoverview mediaUtils - Utilities for media source detection, Base64 decoding, and MIME mapping.
 */

'use strict'

const fsExtra = require('fs-extra')
const { PPTXError } = require('./errors.js')

/**
 * Common media extension to MIME type mapping.
 */
const EXT_TO_MIME = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  svg: 'image/svg+xml',
  tiff: 'image/tiff',
  tif: 'image/tiff',
  bmp: 'image/bmp',
  wmf: 'image/x-wmf',
  emf: 'image/x-emf',
  webp: 'image/webp',
  mp4: 'video/mp4',
  avi: 'video/x-msvideo',
  mov: 'video/quicktime',
  wmv: 'video/x-ms-wmv',
  m4v: 'video/x-m4v',
  webm: 'video/webm',
  mp3: 'audio/mpeg',
  wav: 'audio/wav',
  m4a: 'audio/mp4',
  wma: 'audio/x-ms-wma',
  ogg: 'audio/ogg',
  aac: 'audio/aac',
}

/**
 * Reverse mapping from MIME type to preferred extension.
 */
const MIME_TO_EXT = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/jpg': 'jpg',
  'image/gif': 'gif',
  'image/svg+xml': 'svg',
  'image/tiff': 'tiff',
  'image/bmp': 'bmp',
  'image/webp': 'webp',
  'image/x-wmf': 'wmf',
  'image/x-emf': 'emf',
  'video/mp4': 'mp4',
  'video/x-msvideo': 'avi',
  'video/quicktime': 'mov',
  'video/x-ms-wmv': 'wmv',
  'video/x-m4v': 'm4v',
  'video/webm': 'webm',
  'audio/mpeg': 'mp3',
  'audio/mp3': 'mp3',
  'audio/wav': 'wav',
  'audio/x-wav': 'wav',
  'audio/mp4': 'm4a',
  'audio/x-ms-wma': 'wma',
  'audio/ogg': 'ogg',
  'audio/aac': 'aac',
}

/**
 * 1x1 transparent PNG fallback poster.
 */
const DEFAULT_POSTER_PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
  'base64'
)

/**
 * Opaque 16:9 dark slate PNG thumbnail with play button.
 * Eliminates the Windows DirectShow green chroma-key placeholder in PowerPoint.
 */
const DEFAULT_VIDEO_POSTER_PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAUAAAAC0CAIAAABqhmJGAAAC+ElEQVR42u3cS1KTURSFUWagiEl4JAgChpcPQEWgxSCY/1Bix4b2pPLf5O6TlVojOLu+Sv0hZOvw4gkIteUEIGBAwICAYVMCnp0/AqEEDAIGBAwIGAQMCBgQMCBgKBbwdP4AhBIwCBgQMCBgEDAgYEDAgIChWMAHn34BoQQMAgYEDAgYBAwIGBAwIGAoFvD+2T0QSsAgYEDAgIBBwICAAQEDAoZiAe+d/gRCCRgEDAgYEDAIGBAwIGBAwFAs4N2TH0AoAYOAAQEDAgYBAwIGBAwIGIoFPPn4HQglYBAwIGAGd/f84ggCJqPV/+dcAiavWyULmCLpyljAxKcrYwFToV4NRwY8Pr4jRdN0/+bUKQSsXg0LmFr1aljAZNerYQGTXa+GBUx2vRoWMAJGwOrVMAJWr4ZLBTw6uqVDHQZslA4JWL0aFjACRsAsWe9isdCwgAkOuHXDBhIwbQNumrGBBEyrp9/Fvy8NC5jggBtlbCYBs7qAB8/YTB0F/P7DDf1oF/CADZupHwLeoICHythMAmZtAS/fsJkEzDoDXjJjMwmY4et9bcDLZGwsAeMdGAEL2DOwgPEptIAFjL8Ds9qAdw6/0Y+Ib2KZqR8Crh/w4N+FNpOAWUXAjf4byUwCpknD6hUw8QE3/UUOAwmYhgG3/k0sAwmYtk/C6hUwAhawgCnUsFF6DPjd7Csd6jBgo3RIwBpWr4Ap3bAhBIyAEbCG1YuANaxeAVO8YWcXMKkNO7iASW3YqWMC3p5+IUvTdJ03i4A1rF4BE56xYwqYyIwdUMDklexcAiajakcQMCBgQMDAn4DfHnwGQgkYBAwIGBAwCBgQMCBgQMBQLOA3+9dAKAGDgAEBAwIGAQMCBgQMCBiqBbx3BYQSMAgYEDAgYBAwIGBAwICAoVrAu5dAKAGDgAEBAwIGAQMCBgQMCBiqBTy5AEIJGAQMCBgQMAgYEDAgYEDAUC3g8TkQSsAgYEDAgIBBwICAAQEDAoZqAY/mQCgBg4ABAQOv8hugaN1hxlx7XgAAAABJRU5ErkJggg==',
  'base64'
)

let defaultVideoPoster = null

/**
 * Returns default video poster PNG buffer (16:9 opaque canvas with play symbol).
 * Eliminates green box overlay in PowerPoint when no poster is provided.
 *
 * @returns {Buffer}
 */
function getDefaultVideoPoster() {
  if (!defaultVideoPoster) {
    try {
      const path = require('path')
      const posterPath = path.join(__dirname, '../templates/default_video_poster.png')
      if (fsExtra.existsSync(posterPath)) {
        defaultVideoPoster = fsExtra.readFileSync(posterPath)
      }
    } catch {
      // fallback
    }
    if (!defaultVideoPoster) {
      defaultVideoPoster = DEFAULT_VIDEO_POSTER_PNG
    }
  }
  return defaultVideoPoster
}

let defaultAudioIcon = null

/**
 * Returns default audio speaker icon PNG buffer.
 *
 * @returns {Buffer}
 */
function getDefaultAudioIcon() {
  if (!defaultAudioIcon) {
    try {
      const path = require('path')
      const iconPath = path.join(__dirname, '../templates/speaker_icon.png')
      if (fsExtra.existsSync(iconPath)) {
        defaultAudioIcon = fsExtra.readFileSync(iconPath)
      }
    } catch {
      // fallback
    }
    if (!defaultAudioIcon) {
      defaultAudioIcon = DEFAULT_POSTER_PNG
    }
  }
  return defaultAudioIcon
}

/**
 * Checks whether a string represents a Base64 Data URI.
 *
 * @param {string} str
 * @returns {boolean}
 */
function isDataUri(str) {
  return typeof str === 'string' && /^data:[a-zA-Z0-9+/-]+;base64,/i.test(str.trim())
}

/**
 * Checks whether a string is likely a raw Base64-encoded string.
 *
 * @param {string} str
 * @returns {boolean}
 */
function isRawBase64(str) {
  if (typeof str !== 'string') return false
  const trimmed = str.replace(/[\r\n\s]+/g, '')
  if (trimmed.length < 16 || trimmed.length % 4 !== 0) return false
  return /^[A-Za-z0-9+/]+={0,2}$/.test(trimmed)
}

/**
 * Detects file extension from binary payload header/magic bytes.
 *
 * @param {Buffer} data
 * @returns {string}
 */
function detectMediaExtension(data) {
  if (!data || data.length < 4) return 'png'

  const sig = data.slice(0, 32)

  // PNG: 89 50 4E 47
  if (sig[0] === 0x89 && sig[1] === 0x50 && sig[2] === 0x4e && sig[3] === 0x47) return 'png'
  // JPEG: FF D8 FF
  if (sig[0] === 0xff && sig[1] === 0xd8 && sig[2] === 0xff) return 'jpg'
  // GIF: 47 49 46
  if (sig[0] === 0x47 && sig[1] === 0x49 && sig[2] === 0x46) return 'gif'
  // WEBP / WAV / AVI (RIFF...)
  if (sig[0] === 0x52 && sig[1] === 0x49 && sig[2] === 0x46 && sig[3] === 0x46) {
    if (
      sig.length >= 12 &&
      sig[8] === 0x57 &&
      sig[9] === 0x45 &&
      sig[10] === 0x42 &&
      sig[11] === 0x50
    ) {
      return 'webp'
    }
    if (
      sig.length >= 12 &&
      sig[8] === 0x57 &&
      sig[9] === 0x41 &&
      sig[10] === 0x56 &&
      sig[11] === 0x45
    ) {
      return 'wav'
    }
    if (sig.length >= 12 && sig[8] === 0x41 && sig[9] === 0x56 && sig[10] === 0x49) {
      return 'avi'
    }
  }
  // TIFF: II*. or MM.*
  if (
    (sig[0] === 0x49 && sig[1] === 0x49 && sig[2] === 0x2a) ||
    (sig[0] === 0x4d && sig[1] === 0x4d && sig[3] === 0x2a)
  ) {
    return 'tiff'
  }
  // BMP: 42 4D
  if (sig[0] === 0x42 && sig[1] === 0x4d) return 'bmp'
  // MP3: ID3 or frame sync (FF FB/F3/F2)
  if (sig[0] === 0x49 && sig[1] === 0x44 && sig[2] === 0x33) return 'mp3'
  if (sig[0] === 0xff && (sig[1] & 0xe0) === 0xe0) return 'mp3'
  // MP4 / MOV: ....ftyp or ....moov
  if (
    sig.length >= 8 &&
    (sig.slice(4, 8).toString('latin1') === 'ftyp' || sig.slice(4, 8).toString('latin1') === 'moov')
  ) {
    return 'mp4'
  }
  // OGG: OggS
  if (sig[0] === 0x4f && sig[1] === 0x67 && sig[2] === 0x67 && sig[3] === 0x53) return 'ogg'

  // SVG text
  const textStart = data.slice(0, 100).toString('utf8').trim()
  if (
    textStart.startsWith('<svg') ||
    textStart.includes('<svg') ||
    (textStart.startsWith('<?xml') && textStart.includes('<svg'))
  ) {
    return 'svg'
  }

  return 'png'
}

/**
 * Decodes a Data URI or raw Base64 string into a Buffer.
 *
 * @param {string} source
 * @returns {{ data: Buffer, mimeType?: string, ext?: string }}
 */
function decodeBase64Source(source) {
  const trimmed = source.trim()
  if (isDataUri(trimmed)) {
    const commaIdx = trimmed.indexOf(',')
    const header = trimmed.slice(0, commaIdx)
    const base64Data = trimmed.slice(commaIdx + 1).replace(/[\r\n\s]+/g, '')
    const mimeMatch = header.match(/^data:([^;,]+)/i)
    const mimeType = mimeMatch ? mimeMatch[1].toLowerCase() : undefined
    const data = Buffer.from(base64Data, 'base64')
    const ext = (mimeType && MIME_TO_EXT[mimeType]) || detectMediaExtension(data)
    return { data, mimeType: mimeType || EXT_TO_MIME[ext], ext }
  }

  // Raw base64
  const cleaned = trimmed.replace(/[\r\n\s]+/g, '')
  const data = Buffer.from(cleaned, 'base64')
  const ext = detectMediaExtension(data)
  const mimeType = EXT_TO_MIME[ext] || 'application/octet-stream'
  return { data, mimeType, ext }
}

/**
 * Normalizes any media source (file path, URL, data URI, raw base64, Buffer, or stream) into a Buffer and metadata.
 *
 * @param {string|Buffer|Uint8Array|Object} source
 * @param {string} [defaultCategory='image'] - 'image' | 'video' | 'audio'
 * @param {string} [explicitMimeType]
 * @returns {Promise<{ data: Buffer, ext: string, mimeType: string }>}
 */
async function resolveMediaPayload(source, defaultCategory = 'image', explicitMimeType) {
  if (!source) {
    throw new PPTXError('Media source cannot be null or undefined')
  }

  // 1. ReadableStream
  if (
    typeof source === 'object' &&
    typeof source.on === 'function' &&
    typeof source.pipe === 'function'
  ) {
    const chunks = []
    await new Promise((resolve, reject) => {
      source.on('data', chunk => chunks.push(chunk))
      source.on('end', resolve)
      source.on('error', reject)
    })
    const data = Buffer.concat(chunks)
    const ext =
      source.path && typeof source.path === 'string'
        ? source.path.split('.').pop().toLowerCase()
        : detectMediaExtension(data)
    const mimeType =
      explicitMimeType ||
      EXT_TO_MIME[ext] ||
      (defaultCategory === 'video'
        ? 'video/mp4'
        : defaultCategory === 'audio'
          ? 'audio/mpeg'
          : 'image/png')
    return { data, ext, mimeType }
  }

  // 2. Buffer or Uint8Array
  if (Buffer.isBuffer(source) || source instanceof Uint8Array) {
    const data = Buffer.from(source)
    const ext = detectMediaExtension(data)
    const mimeType =
      explicitMimeType ||
      EXT_TO_MIME[ext] ||
      (defaultCategory === 'video'
        ? 'video/mp4'
        : defaultCategory === 'audio'
          ? 'audio/mpeg'
          : 'image/png')
    return { data, ext, mimeType }
  }

  // 3. String
  if (typeof source === 'string') {
    const trimmed = source.trim()

    // 3.a Remote URL
    if (trimmed.startsWith('http://') || trimmed.startsWith('https://')) {
      const res = await fetch(trimmed)
      if (!res.ok) {
        throw new PPTXError(
          `Failed to fetch remote media from "${trimmed}": HTTP ${res.status} ${res.statusText}`
        )
      }
      const ab = await res.arrayBuffer()
      const data = Buffer.from(ab)
      const urlPath = trimmed.split('?')[0]
      const urlExt = urlPath.includes('.') ? urlPath.split('.').pop().toLowerCase() : ''
      const ext = EXT_TO_MIME[urlExt] ? urlExt : detectMediaExtension(data)
      const headerMime = res.headers.get('content-type')
      const mimeType =
        explicitMimeType ||
        headerMime ||
        EXT_TO_MIME[ext] ||
        (defaultCategory === 'video'
          ? 'video/mp4'
          : defaultCategory === 'audio'
            ? 'audio/mpeg'
            : 'image/png')
      return { data, ext, mimeType }
    }

    // 3.b Data URI
    if (isDataUri(trimmed)) {
      const decoded = decodeBase64Source(trimmed)
      const mimeType =
        explicitMimeType ||
        decoded.mimeType ||
        (defaultCategory === 'video'
          ? 'video/mp4'
          : defaultCategory === 'audio'
            ? 'audio/mpeg'
            : 'image/png')
      return { data: decoded.data, ext: decoded.ext, mimeType }
    }

    // 3.c Check if file exists on disk
    let existsOnDisk = false
    try {
      existsOnDisk = fsExtra.existsSync(trimmed)
    } catch {
      existsOnDisk = false
    }

    if (existsOnDisk) {
      const data = await fsExtra.readFile(trimmed)
      const fileExt = trimmed.includes('.') ? trimmed.split('.').pop().toLowerCase() : ''
      const ext = EXT_TO_MIME[fileExt] ? fileExt : detectMediaExtension(data)
      const mimeType =
        explicitMimeType ||
        EXT_TO_MIME[ext] ||
        (defaultCategory === 'video'
          ? 'video/mp4'
          : defaultCategory === 'audio'
            ? 'audio/mpeg'
            : 'image/png')
      return { data, ext, mimeType }
    }

    // 3.d Raw Base64 string
    if (isRawBase64(trimmed)) {
      const decoded = decodeBase64Source(trimmed)
      const mimeType =
        explicitMimeType ||
        decoded.mimeType ||
        (defaultCategory === 'video'
          ? 'video/mp4'
          : defaultCategory === 'audio'
            ? 'audio/mpeg'
            : 'image/png')
      return { data: decoded.data, ext: decoded.ext, mimeType }
    }

    // Fallback: attempt to read file, or throw clear error
    try {
      const data = await fsExtra.readFile(trimmed)
      const fileExt = trimmed.includes('.') ? trimmed.split('.').pop().toLowerCase() : ''
      const ext = EXT_TO_MIME[fileExt] ? fileExt : detectMediaExtension(data)
      const mimeType =
        explicitMimeType ||
        EXT_TO_MIME[ext] ||
        (defaultCategory === 'video'
          ? 'video/mp4'
          : defaultCategory === 'audio'
            ? 'audio/mpeg'
            : 'image/png')
      return { data, ext, mimeType }
    } catch (err) {
      throw new PPTXError(
        `Cannot resolve media source "${trimmed.substring(0, 60)}...": file not found and not valid Base64 data.`,
        err
      )
    }
  }

  throw new PPTXError(
    'Unsupported media source type. Expected file path, URL, Base64 string, or Buffer.'
  )
}

module.exports = {
  EXT_TO_MIME,
  MIME_TO_EXT,
  DEFAULT_POSTER_PNG,
  DEFAULT_VIDEO_POSTER_PNG,
  getDefaultVideoPoster,
  getDefaultAudioIcon,
  isDataUri,
  isRawBase64,
  detectMediaExtension,
  decodeBase64Source,
  resolveMediaPayload,
}
