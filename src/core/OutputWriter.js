/**
 * @fileoverview OutputWriter bridge — re-exports from src/writer/OutputWriter.js.
 * Preserves backwards compatibility for existing internal imports.
 *
 * @module core/OutputWriter
 */

'use strict'

const { OutputWriter, formatXml } = require('../writer/OutputWriter.js')

module.exports = { OutputWriter, formatXml }
