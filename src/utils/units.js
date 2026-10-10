/**
 * @fileoverview Units - Spatial Units Utility Module
 * Converts points (pt), inches (in), and pixels (px) to integer English Metric Units (EMUs).
 *
 * OpenXML English Metric Units (EMUs) standards:
 * - 1 pt = 12,700 EMUs
 * - 1 in = 914,400 EMUs
 * - 1 px (96 DPI) = 9,525 EMUs
 */

'use strict'

const EMU_PER_PT = 12700
const EMU_PER_IN = 914400
const EMU_PER_PX = 9525

function ptToEmu(pt) {
  return Math.round(pt * EMU_PER_PT)
}

function inToEmu(inches) {
  return Math.round(inches * EMU_PER_IN)
}

function pxToEmu(px) {
  return Math.round(px * EMU_PER_PX)
}

function emuToPt(emu) {
  return emu / EMU_PER_PT
}

function emuToIn(emu) {
  return emu / EMU_PER_IN
}

function emuToPx(emu) {
  return emu / EMU_PER_PX
}

/**
 * Normalizes a spatial coordinate or dimension into EMUs.
 * Handles strings (e.g., '10pt', '2in', '100px') and numbers.
 * @param {string|number} value
 * @returns {number} Integer EMUs
 */
function toEmu(value) {
  if (value === undefined || value === null) return 0

  if (typeof value === 'number') {
    // If it's a small number, assume inches (legacy compatibility behavior)
    if (value < 10) return inToEmu(value)
    // If it's slightly larger, assume pt
    if (value <= 10000) return ptToEmu(value)
    // Otherwise assume it's already EMUs
    return Math.round(value)
  }

  if (typeof value === 'string') {
    const val = value.trim().toLowerCase()
    const num = parseFloat(val)
    if (isNaN(num)) return 0

    if (val.endsWith('in')) return inToEmu(num)
    if (val.endsWith('pt')) return ptToEmu(num)
    if (val.endsWith('px')) return pxToEmu(num)
    if (val.endsWith('emu')) return Math.round(num)

    // Default fallback parsing identical to number behavior
    if (num < 10) return inToEmu(num)
    if (num <= 10000) return ptToEmu(num)
    return Math.round(num)
  }

  return 0
}

module.exports = {
  EMU_PER_PT,
  EMU_PER_IN,
  EMU_PER_PX,
  ptToEmu,
  inToEmu,
  pxToEmu,
  emuToPt,
  emuToIn,
  emuToPx,
  toEmu
}
