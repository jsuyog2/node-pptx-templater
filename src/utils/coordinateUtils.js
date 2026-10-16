/**
 * @fileoverview coordinateUtils.js - Coordinate & Spatial Resolution Utilities
 * Handles percentage-based coordinates (0–100%), EMU calculations, and canvas dimension mapping.
 */

'use strict'

const { inToEmu, ptToEmu, pxToEmu } = require('../spatial/units.js')

const DEFAULT_SLIDE_WIDTH_EMU = 12192000 // 16:9 widescreen default (13.333 inches)
const DEFAULT_SLIDE_HEIGHT_EMU = 6858000 // 16:9 widescreen default (7.5 inches)

/**
 * Resolves presentation slide dimensions in EMUs from slideManager or defaults.
 *
 * @param {Object} [slideManager] - SlideManager instance.
 * @returns {{ width: number, height: number }}
 */
function getCanvasDimensions(slideManager) {
  if (slideManager && typeof slideManager.getSlideDimensions === 'function') {
    const dims = slideManager.getSlideDimensions()
    if (dims && dims.width && dims.height) {
      return { width: dims.width, height: dims.height }
    }
  }
  return { width: DEFAULT_SLIDE_WIDTH_EMU, height: DEFAULT_SLIDE_HEIGHT_EMU }
}

/**
 * Resolves a single spatial coordinate or dimension into EMUs.
 * Supports:
 * - String percentages: '0%', '25%', '50%', '100%'
 * - Numeric percentages: 0 to 100 when options.unit is '%' or 'percent', or options.isPercent is true
 * - Dimension strings: '1.5in', '72pt', '100px', '952500emu'
 * - Numbers: > 50,000 treated as EMUs, <= 50,000 treated as pixels by default
 *
 * @param {string|number} value - The input coordinate or dimension.
 * @param {number} totalCanvasEmu - The total dimension in EMUs (canvas width or height).
 * @param {Object} [options={}] - Options specifying units or percentage modes.
 * @returns {number} Value in English Metric Units (EMUs).
 */
function resolveCoordinate(value, totalCanvasEmu = DEFAULT_SLIDE_WIDTH_EMU, options = {}) {
  if (value === undefined || value === null) {
    return 0
  }

  // 1. String handling
  if (typeof value === 'string') {
    const trimmed = value.trim()

    // Percentage string: e.g. "50%", "100%", "33.3%"
    if (trimmed.endsWith('%')) {
      const pct = parseFloat(trimmed)
      if (!isNaN(pct)) {
        return Math.round((pct / 100) * totalCanvasEmu)
      }
      return 0
    }

    const lower = trimmed.toLowerCase()
    const num = parseFloat(lower)
    if (isNaN(num)) return 0

    if (lower.endsWith('in')) return inToEmu(num)
    if (lower.endsWith('pt')) return ptToEmu(num)
    if (lower.endsWith('px')) return pxToEmu(num)
    if (lower.endsWith('emu')) return Math.round(num)

    // String number without unit suffix: check options
    const unit = options.unit ? String(options.unit).toLowerCase() : null
    if (unit === '%' || unit === 'percent' || options.isPercent || options.usePercent) {
      return Math.round((num / 100) * totalCanvasEmu)
    }
    if (unit === 'in') return inToEmu(num)
    if (unit === 'pt') return ptToEmu(num)
    if (unit === 'emu') return Math.round(num)
    if (unit === 'px') return pxToEmu(num)

    if (num > 50000) return Math.round(num)
    return pxToEmu(num)
  }

  // 2. Numeric handling
  if (typeof value === 'number') {
    const unit = options.unit ? String(options.unit).toLowerCase() : null
    if (unit === '%' || unit === 'percent' || options.isPercent || options.usePercent) {
      return Math.round((value / 100) * totalCanvasEmu)
    }
    if (unit === 'in') return inToEmu(value)
    if (unit === 'pt') return ptToEmu(value)
    if (unit === 'emu') return Math.round(value)
    if (unit === 'px') return pxToEmu(value)

    if (value > 50000) return Math.round(value)
    return pxToEmu(value)
  }

  return 0
}

/**
 * Resolves full element bounding box (x, y, width, height) in EMUs against slide dimensions.
 *
 * @param {Object} options - Options containing x, y, width, height (or left, top, w, h).
 * @param {Object} [slideManager] - SlideManager instance for dimension discovery.
 * @returns {{ x: number, y: number, width: number, height: number }} Resolved bounds in EMUs.
 */
function resolveElementBounds(options = {}, slideManager = null) {
  const dims = getCanvasDimensions(slideManager)
  const canvasWidth = dims.width
  const canvasHeight = dims.height

  const rawX = options.x !== undefined ? options.x : options.left !== undefined ? options.left : 0
  const rawY = options.y !== undefined ? options.y : options.top !== undefined ? options.top : 0
  const rawW =
    options.width !== undefined
      ? options.width
      : options.w !== undefined
        ? options.w
        : options.size !== undefined
          ? options.size
          : 100
  const rawH =
    options.height !== undefined
      ? options.height
      : options.h !== undefined
        ? options.h
        : options.size !== undefined
          ? options.size
          : 100

  return {
    x: resolveCoordinate(rawX, canvasWidth, options),
    y: resolveCoordinate(rawY, canvasHeight, options),
    width: resolveCoordinate(rawW, canvasWidth, options),
    height: resolveCoordinate(rawH, canvasHeight, options),
  }
}

module.exports = {
  DEFAULT_SLIDE_WIDTH_EMU,
  DEFAULT_SLIDE_HEIGHT_EMU,
  getCanvasDimensions,
  resolveCoordinate,
  resolveElementBounds,
}
