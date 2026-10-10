/**
 * @fileoverview units.js - OpenXML English Metric Units (EMU) and Spatial Bounding Box Engine
 *
 * OpenXML Standards:
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

function toEmu(value) {
  if (value === undefined || value === null) return 0

  if (typeof value === 'number') {
    if (value < 10) return inToEmu(value)
    if (value <= 10000) return ptToEmu(value)
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

    if (num < 10) return inToEmu(num)
    if (num <= 10000) return ptToEmu(num)
    return Math.round(num)
  }

  return 0
}

function checkCollision(boxA, boxB) {
  return (
    boxA.x < boxB.x + boxB.width &&
    boxA.x + boxA.width > boxB.x &&
    boxA.y < boxB.y + boxB.height &&
    boxA.y + boxA.height > boxB.y
  )
}

function isContained(inner, outer) {
  return (
    inner.x >= outer.x &&
    inner.y >= outer.y &&
    inner.x + inner.width <= outer.x + outer.width &&
    inner.y + inner.height <= outer.y + outer.height
  )
}

function getBounds(element) {
  const xfrm =
    element?.['p:spPr']?.['a:xfrm'] ||
    element?.['p:picPr']?.['a:xfrm'] ||
    element?.['p:xfrm'] ||
    element?.['a:xfrm']

  const off = xfrm?.['a:off']
  const ext = xfrm?.['a:ext']

  return {
    x: parseInt(off?.['@_x'] || '0', 10),
    y: parseInt(off?.['@_y'] || '0', 10),
    width: parseInt(ext?.['@_cx'] || '0', 10),
    height: parseInt(ext?.['@_cy'] || '0', 10),
  }
}

function setBounds(element, bounds) {
  const xfrm =
    element?.['p:spPr']?.['a:xfrm'] ||
    element?.['p:picPr']?.['a:xfrm'] ||
    element?.['p:xfrm'] ||
    element?.['a:xfrm']

  if (!xfrm) return

  if (!xfrm['a:off']) xfrm['a:off'] = {}
  if (!xfrm['a:ext']) xfrm['a:ext'] = {}

  if (bounds.x !== undefined) xfrm['a:off']['@_x'] = String(Math.round(bounds.x))
  if (bounds.y !== undefined) xfrm['a:off']['@_y'] = String(Math.round(bounds.y))
  if (bounds.width !== undefined) xfrm['a:ext']['@_cx'] = String(Math.round(bounds.width))
  if (bounds.height !== undefined) xfrm['a:ext']['@_cy'] = String(Math.round(bounds.height))
}

function move(element, dx, dy) {
  const current = getBounds(element)
  setBounds(element, {
    x: current.x + Math.round(dx),
    y: current.y + Math.round(dy),
  })
}

function resize(element, dw, dh) {
  const current = getBounds(element)
  setBounds(element, {
    width: Math.max(0, current.width + Math.round(dw)),
    height: Math.max(0, current.height + Math.round(dh)),
  })
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
  toEmu,
  checkCollision,
  isContained,
  getBounds,
  setBounds,
  move,
  resize,
}
