/**
 * @fileoverview units.ts - OpenXML English Metric Units (EMU) and Spatial Bounding Box Engine
 *
 * OpenXML Standards:
 * - 1 pt = 12,700 EMUs
 * - 1 in = 914,400 EMUs
 * - 1 px (96 DPI) = 9,525 EMUs
 */

export const EMU_PER_PT = 12700
export const EMU_PER_IN = 914400
export const EMU_PER_PX = 9525

export interface BoundingBox {
  x: number
  y: number
  width: number
  height: number
}

export function ptToEmu(pt: number): number {
  return Math.round(pt * EMU_PER_PT)
}

export function inToEmu(inches: number): number {
  return Math.round(inches * EMU_PER_IN)
}

export function pxToEmu(px: number): number {
  return Math.round(px * EMU_PER_PX)
}

export function emuToPt(emu: number): number {
  return emu / EMU_PER_PT
}

export function emuToIn(emu: number): number {
  return emu / EMU_PER_IN
}

export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX
}

/**
 * Normalizes a spatial coordinate or dimension into integer EMUs.
 */
export function toEmu(value: string | number | undefined | null): number {
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

/**
 * Tests for AABB collision/overlap between two bounding boxes.
 */
export function checkCollision(boxA: BoundingBox, boxB: BoundingBox): boolean {
  return (
    boxA.x < boxB.x + boxB.width &&
    boxA.x + boxA.width > boxB.x &&
    boxA.y < boxB.y + boxB.height &&
    boxA.y + boxA.height > boxB.y
  )
}

/**
 * Tests if boxA is completely contained inside boxB.
 */
export function isContained(inner: BoundingBox, outer: BoundingBox): boolean {
  return (
    inner.x >= outer.x &&
    inner.y >= outer.y &&
    inner.x + inner.width <= outer.x + outer.width &&
    inner.y + inner.height <= outer.y + outer.height
  )
}

/**
 * Extracts bounding box (x, y, width, height) in EMUs from an OpenXML shape node (p:sp, p:pic, p:graphicFrame).
 */
export function getBounds(element: any): BoundingBox {
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

/**
 * Updates bounding box attributes on an OpenXML shape node.
 */
export function setBounds(element: any, bounds: Partial<BoundingBox>): void {
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

/**
 * Moves an element by dx and dy EMUs.
 */
export function move(element: any, dx: number, dy: number): void {
  const current = getBounds(element)
  setBounds(element, {
    x: current.x + Math.round(dx),
    y: current.y + Math.round(dy),
  })
}

/**
 * Resizes an element by delta width and delta height EMUs.
 */
export function resize(element: any, dw: number, dh: number): void {
  const current = getBounds(element)
  setBounds(element, {
    width: Math.max(0, current.width + Math.round(dw)),
    height: Math.max(0, current.height + Math.round(dh)),
  })
}
