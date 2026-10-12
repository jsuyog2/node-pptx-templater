/**
 * @fileoverview Web framework and runtime response adapters for PPTX streaming.
 *
 * Provides native adapters for Express.js, Fastify, Koa, and the standard Web Fetch API Response
 * (Next.js Route Handlers, Remix, SvelteKit, Cloudflare Workers, Vercel Edge).
 *
 * @module FrameworkAdapters
 */

'use strict'

const PPTX_MIME_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation'

/**
 * Sends a PPTX presentation via an Express.js response.
 *
 * @param {Object} res - Express response object.
 * @param {import('stream').Readable|Buffer} streamOrBuffer - Stream or buffer to send.
 * @param {string} [filename='presentation.pptx'] - Download filename.
 * @param {Object} [options={}] - Additional response options.
 */
function sendExpress(res, streamOrBuffer, filename = 'presentation.pptx', options = {}) {
  res.setHeader('Content-Type', PPTX_MIME_TYPE)
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`)

  if (options.headers && typeof options.headers === 'object') {
    for (const [key, value] of Object.entries(options.headers)) {
      res.setHeader(key, value)
    }
  }

  if (Buffer.isBuffer(streamOrBuffer)) {
    res.setHeader('Content-Length', String(streamOrBuffer.length))
    return res.end(streamOrBuffer)
  }

  if (streamOrBuffer && typeof streamOrBuffer.pipe === 'function') {
    return streamOrBuffer.pipe(res)
  }

  throw new TypeError(
    'sendExpress: Expected Node.js Buffer or Readable stream.'
  )
}

/**
 * Sends a PPTX presentation via a Fastify reply object.
 *
 * @param {Object} reply - Fastify reply object.
 * @param {import('stream').Readable|Buffer} streamOrBuffer - Stream or buffer to send.
 * @param {string} [filename='presentation.pptx'] - Download filename.
 * @param {Object} [options={}] - Additional reply options.
 */
function sendFastify(reply, streamOrBuffer, filename = 'presentation.pptx', options = {}) {
  reply.header('Content-Type', PPTX_MIME_TYPE)
  reply.header('Content-Disposition', `attachment; filename="${filename}"`)

  if (options.headers && typeof options.headers === 'object') {
    for (const [key, value] of Object.entries(options.headers)) {
      reply.header(key, value)
    }
  }

  return reply.send(streamOrBuffer)
}

/**
 * Sends a PPTX presentation via a Koa context object.
 *
 * @param {Object} ctx - Koa context object.
 * @param {import('stream').Readable|Buffer} streamOrBuffer - Stream or buffer to send.
 * @param {string} [filename='presentation.pptx'] - Download filename.
 * @param {Object} [options={}] - Additional context options.
 */
function sendKoa(ctx, streamOrBuffer, filename = 'presentation.pptx', options = {}) {
  ctx.type = PPTX_MIME_TYPE
  ctx.set('Content-Disposition', `attachment; filename="${filename}"`)

  if (options.headers && typeof options.headers === 'object') {
    for (const [key, value] of Object.entries(options.headers)) {
      ctx.set(key, value)
    }
  }

  ctx.body = streamOrBuffer
}

/**
 * Emits a standard Web API Response object (Fetch API / Edge runtime compliant).
 *
 * @param {Buffer|ArrayBuffer|Uint8Array} bufferOrArrayBuffer - Binary representation.
 * @param {string} [filename='presentation.pptx'] - Filename in Content-Disposition.
 * @param {Object} [options={}] - Response options (status, headers).
 * @returns {Response}
 */
function createWebResponse(
  bufferOrArrayBuffer,
  filename = 'presentation.pptx',
  options = {}
) {
  let arrayBuffer
  let byteLength

  if (Buffer.isBuffer(bufferOrArrayBuffer)) {
    arrayBuffer = bufferOrArrayBuffer.buffer.slice(
      bufferOrArrayBuffer.byteOffset,
      bufferOrArrayBuffer.byteOffset + bufferOrArrayBuffer.byteLength
    )
    byteLength = bufferOrArrayBuffer.length
  } else if (bufferOrArrayBuffer instanceof Uint8Array) {
    arrayBuffer = bufferOrArrayBuffer.buffer.slice(
      bufferOrArrayBuffer.byteOffset,
      bufferOrArrayBuffer.byteOffset + bufferOrArrayBuffer.byteLength
    )
    byteLength = bufferOrArrayBuffer.byteLength
  } else if (bufferOrArrayBuffer instanceof ArrayBuffer) {
    arrayBuffer = bufferOrArrayBuffer
    byteLength = bufferOrArrayBuffer.byteLength
  } else {
    throw new TypeError(
      'createWebResponse: expected Buffer, Uint8Array, or ArrayBuffer.'
    )
  }

  const headers = new Headers({
    'Content-Type': PPTX_MIME_TYPE,
    'Content-Disposition': `attachment; filename="${filename}"`,
    'Content-Length': String(byteLength),
    ...(options.headers || {}),
  })

  return new Response(arrayBuffer, {
    status: options.status || 200,
    headers,
  })
}

module.exports = {
  PPTX_MIME_TYPE,
  sendExpress,
  sendFastify,
  sendKoa,
  createWebResponse,
}
