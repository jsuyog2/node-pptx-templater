/**
 * @fileoverview StyleBuilder - DrawingML OpenXML typography and styling engine.
 * Maps 31 style properties into DrawingML <a:pPr>, <a:rPr>, and <a:bodyPr> structures.
 */

'use strict'

const { EMU_PER_PT } = require('../spatial/units.js')

const PT_EMU = EMU_PER_PT || 12700

/**
 * Normalizes a hex color string to RRGGBB uppercase.
 * @param {string} hex
 * @returns {string}
 */
function normalizeHexColor(hex) {
  if (!hex || typeof hex !== 'string') return '000000'
  const clean = hex.replace('#', '').trim().toUpperCase()
  if (clean.length === 3) {
    return clean
      .split('')
      .map(c => c + c)
      .join('')
  }
  return clean.padStart(6, '0').slice(0, 6)
}

/**
 * Capitalizes the first letter of each word in a string.
 * @param {string} str
 * @returns {string}
 */
function toCapitalized(str) {
  if (!str || typeof str !== 'string') return str
  return str.replace(/\b\w/g, c => c.toUpperCase())
}

/**
 * Applies text casing transformations based on style properties.
 * @param {string} text
 * @param {Object} style
 * @returns {string}
 */
function applyTextCasing(text, style = {}) {
  if (text === null || text === undefined) return text
  let result = String(text)
  if (style.uppercase) {
    result = result.toUpperCase()
  } else if (style.lowercase) {
    result = result.toLowerCase()
  } else if (style.capitalize) {
    result = toCapitalized(result)
  }
  return result
}

class StyleBuilder {
  /**
   * Builds or updates an <a:rPr> (Run Properties) AST node from a style object.
   *
   * Supported Run Properties:
   * - fontSize (pt -> sz in hundredths of a pt)
   * - color (hex -> a:solidFill/a:srgbClr)
   * - highlight (hex -> a:highlight/a:srgbClr)
   * - underline (boolean or string -> u)
   * - bold (boolean -> b)
   * - italic (boolean -> i)
   * - superscript / subscript (baseline: 30000 / -25000)
   * - strikethrough (boolean or string -> strike)
   * - smallCaps (boolean -> cap="small")
   * - uppercase (cap="all" if not transformed textually)
   * - charSpacing (pt -> spc in hundredths of a pt)
   * - shadowColor, textShadow, blur, shadowRadius, shadowOffsetX, shadowOffsetY (a:effectLst/a:outerShdw)
   *
   * @param {Object} existingRPr - Existing <a:rPr> AST object or empty object.
   * @param {Object} style - Style properties to apply.
   * @returns {Object} Updated <a:rPr> object.
   */
  static buildRunProperties(existingRPr = {}, style = {}) {
    const rPr = { ...(existingRPr || {}) }

    // 1. Font Size (points -> hundredths of a point in @_sz)
    if (style.fontSize !== undefined && style.fontSize !== null) {
      rPr['@_sz'] = String(Math.round(Number(style.fontSize) * 100))
    }

    // 2. Text Color
    if (style.color !== undefined && style.color !== null) {
      const hex = normalizeHexColor(style.color)
      rPr['a:solidFill'] = { 'a:srgbClr': { '@_val': hex } }
    }

    // 3. Highlight Color
    if (style.highlight !== undefined && style.highlight !== null) {
      if (style.highlight === false || style.highlight === 'none') {
        delete rPr['a:highlight']
      } else {
        const hex = normalizeHexColor(style.highlight)
        rPr['a:highlight'] = { 'a:srgbClr': { '@_val': hex } }
      }
    }

    // 4. Underline
    if (style.underline !== undefined) {
      if (style.underline === false || style.underline === 'none') {
        delete rPr['@_u']
      } else if (typeof style.underline === 'string') {
        rPr['@_u'] = style.underline
      } else if (style.underline) {
        rPr['@_u'] = 'sng'
      }
    }

    // 5. Bold
    if (style.bold !== undefined) {
      rPr['@_b'] = style.bold ? '1' : '0'
    }

    // 6. Italic
    if (style.italic !== undefined) {
      rPr['@_i'] = style.italic ? '1' : '0'
    }

    // 7. Strikethrough
    if (style.strikethrough !== undefined || style.strike !== undefined) {
      const strikeVal = style.strikethrough !== undefined ? style.strikethrough : style.strike
      if (strikeVal === false || strikeVal === 'none' || strikeVal === 'noStrike') {
        delete rPr['@_strike']
      } else if (strikeVal === 'double' || strikeVal === 'dblStrike') {
        rPr['@_strike'] = 'dblStrike'
      } else if (strikeVal) {
        rPr['@_strike'] = 'sngStrike'
      }
    }

    // 8. Superscript / Subscript
    if (style.superscript) {
      rPr['@_baseline'] = '30000'
    } else if (style.subscript) {
      rPr['@_baseline'] = '-25000'
    } else if (style.superscript === false || style.subscript === false) {
      delete rPr['@_baseline']
    }

    // 9. Casing: smallCaps & uppercase
    if (style.smallCaps !== undefined) {
      if (style.smallCaps) {
        rPr['@_cap'] = 'small'
      } else if (rPr['@_cap'] === 'small') {
        delete rPr['@_cap']
      }
    }
    if (style.uppercase && !style.smallCaps) {
      rPr['@_cap'] = 'all'
    }

    // 10. Character Spacing (points -> hundredths of a point in @_spc)
    if (style.charSpacing !== undefined && style.charSpacing !== null) {
      rPr['@_spc'] = String(Math.round(Number(style.charSpacing) * 100))
    }

    // 11. Shadows & Effects (shadowColor, textShadow, blur, shadowRadius, shadowOffsetX, shadowOffsetY)
    const hasShadow =
      style.textShadow ||
      style.shadowColor !== undefined ||
      style.shadowRadius !== undefined ||
      style.shadowOffsetX !== undefined ||
      style.shadowOffsetY !== undefined

    if (hasShadow) {
      const shadowHex = normalizeHexColor(style.shadowColor || '000000')
      const blurVal =
        style.blur !== undefined
          ? Number(style.blur)
          : style.shadowRadius !== undefined
            ? Number(style.shadowRadius)
            : 3
      const blurRad = Math.round(blurVal * PT_EMU)

      const offsetX = style.shadowOffsetX !== undefined ? Number(style.shadowOffsetX) : 2
      const offsetY = style.shadowOffsetY !== undefined ? Number(style.shadowOffsetY) : 2
      const dist = Math.round(Math.hypot(offsetX, offsetY) * PT_EMU)
      const dirAngle = Math.round(
        (((Math.atan2(offsetY, offsetX) * 180) / Math.PI + 360) % 360) * 60000
      )

      rPr['a:effectLst'] = {
        'a:outerShdw': {
          '@_blurRad': String(blurRad),
          '@_dist': String(dist),
          '@_dir': String(dirAngle),
          '@_algn': 'tl',
          'a:srgbClr': { '@_val': shadowHex },
        },
      }
    } else if (style.textShadow === false) {
      delete rPr['a:effectLst']
    }

    // 12. Font Face / Family
    const font = style.fontFace || style.fontFamily || style.font
    if (font) {
      rPr['a:latin'] = { '@_typeface': String(font) }
      rPr['a:ea'] = { '@_typeface': String(font) }
      rPr['a:cs'] = { '@_typeface': String(font) }
    }

    return rPr
  }

  /**
   * Builds an <a:ln> border element for DrawingML cells or shapes.
   *
   * @param {Object|boolean|string|number} borderOptions - Border configuration.
   * @returns {Object|null} DrawingML line object or null.
   */
  static buildBorder(borderOptions) {
    if (!borderOptions || borderOptions === 'none') {
      return { 'a:noFill': {} }
    }
    const border = typeof borderOptions === 'object' ? borderOptions : {}
    const colorHex = border.color ? normalizeHexColor(border.color) : '000000'
    const ptWidth =
      border.width !== undefined
        ? Number(border.width)
        : typeof borderOptions === 'number'
          ? borderOptions
          : 1
    const emuWidth = Math.round(ptWidth * PT_EMU)
    const dashType = border.dash || border.style || 'solid'

    return {
      '@_w': String(emuWidth),
      '@_cmpd': 'sng',
      'a:solidFill': { 'a:srgbClr': { '@_val': colorHex } },
      'a:prstDash': { '@_val': dashType === 'solid' ? 'solid' : dashType },
    }
  }

  /**
   * Builds or updates an <a:tcPr> (Table Cell Properties) AST node from options.
   *
   * @param {Object} existingTcPr - Existing <a:tcPr> AST object.
   * @param {Object} options - Cell property options.
   * @returns {Object} Updated <a:tcPr> object.
   */
  static buildCellProperties(existingTcPr = {}, options = {}) {
    const tcPr = { ...(existingTcPr || {}) }

    // 1. Cell Solid Fill
    const fill = options.fill !== undefined ? options.fill : options.bgColor || options.background
    if (fill !== undefined) {
      if (fill === false || fill === 'none' || fill === null) {
        delete tcPr['a:solidFill']
        tcPr['a:noFill'] = {}
      } else {
        delete tcPr['a:noFill']
        tcPr['a:solidFill'] = { 'a:srgbClr': { '@_val': normalizeHexColor(fill) } }
      }
    }

    // 2. Vertical Alignment (anchor)
    const vAlign =
      options.verticalAlignment || options.verticalAlign || options.vAlign || options.anchor
    if (vAlign) {
      const v = String(vAlign).toLowerCase()
      if (v === 'top' || v === 't') tcPr['@_anchor'] = 't'
      else if (v === 'middle' || v === 'center' || v === 'ctr') tcPr['@_anchor'] = 'ctr'
      else if (v === 'bottom' || v === 'b') tcPr['@_anchor'] = 'b'
      else if (v === 'dist' || v === 'just') tcPr['@_anchor'] = v
      else tcPr['@_anchor'] = v
    }

    // 3. Cell Margins (marL, marR, marT, marB)
    const margins = options.margins || options.margin
    if (margins !== undefined) {
      if (typeof margins === 'number') {
        const emu = Math.round(margins * PT_EMU)
        tcPr['@_marL'] = String(emu)
        tcPr['@_marR'] = String(emu)
        tcPr['@_marT'] = String(emu)
        tcPr['@_marB'] = String(emu)
      } else if (typeof margins === 'object' && margins !== null) {
        if (margins.left !== undefined)
          tcPr['@_marL'] = String(Math.round(Number(margins.left) * PT_EMU))
        if (margins.right !== undefined)
          tcPr['@_marR'] = String(Math.round(Number(margins.right) * PT_EMU))
        if (margins.top !== undefined)
          tcPr['@_marT'] = String(Math.round(Number(margins.top) * PT_EMU))
        if (margins.bottom !== undefined)
          tcPr['@_marB'] = String(Math.round(Number(margins.bottom) * PT_EMU))
      }
    }
    if (options.marginLeft !== undefined || options.marL !== undefined) {
      tcPr['@_marL'] = String(Math.round(Number(options.marginLeft || options.marL) * PT_EMU))
    }
    if (options.marginRight !== undefined || options.marR !== undefined) {
      tcPr['@_marR'] = String(Math.round(Number(options.marginRight || options.marR) * PT_EMU))
    }
    if (options.marginTop !== undefined || options.marT !== undefined) {
      tcPr['@_marT'] = String(Math.round(Number(options.marginTop || options.marT) * PT_EMU))
    }
    if (options.marginBottom !== undefined || options.marB !== undefined) {
      tcPr['@_marB'] = String(Math.round(Number(options.marginBottom || options.marB) * PT_EMU))
    }

    // 4. Borders (a:lnL, a:lnR, a:lnT, a:lnB)
    if (options.border !== undefined) {
      const ln = StyleBuilder.buildBorder(options.border)
      if (ln) {
        tcPr['a:lnL'] = { ...ln }
        tcPr['a:lnR'] = { ...ln }
        tcPr['a:lnT'] = { ...ln }
        tcPr['a:lnB'] = { ...ln }
      }
    }
    const borderLeft = options.borderLeft !== undefined ? options.borderLeft : options.left
    if (borderLeft !== undefined) {
      const ln = StyleBuilder.buildBorder(borderLeft)
      if (ln) tcPr['a:lnL'] = ln
    }
    const borderRight = options.borderRight !== undefined ? options.borderRight : options.right
    if (borderRight !== undefined) {
      const ln = StyleBuilder.buildBorder(borderRight)
      if (ln) tcPr['a:lnR'] = ln
    }
    const borderTop = options.borderTop !== undefined ? options.borderTop : options.top
    if (borderTop !== undefined) {
      const ln = StyleBuilder.buildBorder(borderTop)
      if (ln) tcPr['a:lnT'] = ln
    }
    const borderBottom = options.borderBottom !== undefined ? options.borderBottom : options.bottom
    if (borderBottom !== undefined) {
      const ln = StyleBuilder.buildBorder(borderBottom)
      if (ln) tcPr['a:lnB'] = ln
    }

    return tcPr
  }

  /**
   * Builds or updates an <a:tblPr> (Table Properties) AST node from options.
   *
   * @param {Object} existingTblPr - Existing <a:tblPr> AST object.
   * @param {Object} options - Table property options.
   * @returns {Object} Updated <a:tblPr> object.
   */
  static buildTableProperties(existingTblPr = {}, options = {}) {
    const tblPr = { ...(existingTblPr || {}) }

    if (options.firstRow !== undefined) tblPr['@_firstRow'] = options.firstRow ? '1' : '0'
    if (options.bandRow !== undefined) tblPr['@_bandRow'] = options.bandRow ? '1' : '0'
    if (options.firstCol !== undefined) tblPr['@_firstCol'] = options.firstCol ? '1' : '0'
    if (options.bandCol !== undefined) tblPr['@_bandCol'] = options.bandCol ? '1' : '0'
    if (options.lastRow !== undefined) tblPr['@_lastRow'] = options.lastRow ? '1' : '0'
    if (options.lastCol !== undefined) tblPr['@_lastCol'] = options.lastCol ? '1' : '0'

    if (options.styleId) {
      let sId = String(options.styleId).trim()
      if (!sId.startsWith('{')) sId = `{${sId}}`
      tblPr['a:tableStyleId'] = sId
    }

    return tblPr
  }

  /**
   * Builds or updates an <a:pPr> (Paragraph Properties) AST node from a style object.
   *
   * Supported Paragraph Properties:
   * - textAlignment / alignment (algn: l, ctr, r, just)
   * - indentLevel / bulletLevel (lvl: 0..8)
   * - lineSpacing (lnSpc in spcPts or spcPct)
   * - bulletChar (a:buChar)
   * - bulletColor (a:buClr/a:srgbClr)
   * - bulletFontSize (a:buSzPts or a:buSzPct)
   * - bulletStyle ('none', 'bullet', 'ordered', etc.)
   *
   * @param {Object} existingPPr - Existing <a:pPr> AST object or empty object.
   * @param {Object} style - Style properties to apply.
   * @returns {Object} Updated <a:pPr> object.
   */
  static buildParagraphProperties(existingPPr = {}, style = {}) {
    const pPr = { ...(existingPPr || {}) }

    // 1. Text Alignment
    const align =
      style.textAlignment !== undefined
        ? style.textAlignment
        : style.alignment !== undefined
          ? style.alignment
          : style.align
    if (align !== undefined && align !== null) {
      const map = {
        left: 'l',
        center: 'ctr',
        centre: 'ctr',
        right: 'r',
        justify: 'just',
        justified: 'just',
        l: 'l',
        ctr: 'ctr',
        r: 'r',
        just: 'just',
      }
      const val = map[String(align).toLowerCase()] || 'l'
      pPr['@_algn'] = val
    }

    // 2. Indent / Bullet Level
    const level = style.indentLevel !== undefined ? style.indentLevel : style.bulletLevel
    if (level !== undefined && level !== null) {
      const lvlNum = Math.max(0, Math.min(8, parseInt(level, 10) || 0))
      pPr['@_lvl'] = String(lvlNum)
    }

    // 3. Line Spacing
    if (style.lineSpacing !== undefined && style.lineSpacing !== null) {
      const spc = style.lineSpacing
      if (typeof spc === 'string' && spc.endsWith('pt')) {
        const pts = parseFloat(spc)
        pPr['a:lnSpc'] = { 'a:spcPts': { '@_val': String(Math.round(pts * 100)) } }
      } else {
        const num = parseFloat(spc)
        if (num > 10) {
          // Points
          pPr['a:lnSpc'] = { 'a:spcPts': { '@_val': String(Math.round(num * 100)) } }
        } else {
          // Percentage (1.0 = 100%, 1.2 = 120%)
          pPr['a:lnSpc'] = { 'a:spcPct': { '@_val': String(Math.round(num * 100000)) } }
        }
      }
    }

    // 4. Bullet Style & Markers
    if (style.bulletStyle === 'none') {
      delete pPr['a:buChar']
      delete pPr['a:buAutoNum']
      pPr['a:buNone'] = {}
    } else if (style.bulletStyle === 'ordered' || style.bulletStyle === 'numeric') {
      delete pPr['a:buNone']
      delete pPr['a:buChar']
      pPr['a:buAutoNum'] = { '@_type': 'arabicPeriod' }
    } else if (style.bulletStyle === 'bullet' || style.bulletChar) {
      delete pPr['a:buNone']
      delete pPr['a:buAutoNum']
      const char = style.bulletChar || '•'
      pPr['a:buChar'] = { '@_char': char }
    }

    // Bullet Color
    if (style.bulletColor !== undefined && style.bulletColor !== null) {
      const hex = normalizeHexColor(style.bulletColor)
      pPr['a:buClr'] = { 'a:srgbClr': { '@_val': hex } }
    }

    // Bullet Font Size
    if (style.bulletFontSize !== undefined && style.bulletFontSize !== null) {
      const sz = Number(style.bulletFontSize)
      if (sz > 20) {
        // Percentage (e.g. 100, 120)
        pPr['a:buSzPct'] = { '@_val': String(Math.round(sz * 1000)) }
      } else {
        // Points
        pPr['a:buSzPts'] = { '@_val': String(Math.round(sz * 100)) }
      }
    }

    return pPr
  }

  /**
   * Builds or updates an <a:bodyPr> (Body Properties) AST node.
   *
   * Supported Body Properties:
   * - verticalAlignment (anchor: t, ctr, b)
   * - textDirection (vert: horz, vert, vert270, wordArtVert)
   *
   * @param {Object} existingBodyPr - Existing <a:bodyPr> AST object.
   * @param {Object} style - Style properties to apply.
   * @returns {Object} Updated <a:bodyPr> object.
   */
  static buildBodyProperties(existingBodyPr = {}, style = {}) {
    const bodyPr = { ...(existingBodyPr || {}) }

    // 1. Vertical Alignment
    if (style.verticalAlignment !== undefined && style.verticalAlignment !== null) {
      const map = {
        top: 't',
        middle: 'ctr',
        center: 'ctr',
        centre: 'ctr',
        bottom: 'b',
        t: 't',
        ctr: 'ctr',
        b: 'b',
      }
      bodyPr['@_anchor'] = map[String(style.verticalAlignment).toLowerCase()] || 't'
    }

    // 2. Text Direction
    if (style.textDirection !== undefined && style.textDirection !== null) {
      const map = {
        horizontal: 'horz',
        horz: 'horz',
        vertical: 'vert',
        vert: 'vert',
        vertical270: 'vert270',
        vert270: 'vert270',
        rotate270: 'vert270',
        wordArtVertical: 'wordArtVert',
        wordArtVert: 'wordArtVert',
      }
      bodyPr['@_vert'] = map[String(style.textDirection)] || 'horz'
    }

    return bodyPr
  }

  /**
   * Removes specified style properties from <a:rPr>, <a:pPr>, and <a:bodyPr>.
   *
   * @param {Object} rPr - Run properties object
   * @param {Object} pPr - Paragraph properties object
   * @param {Object} bodyPr - Body properties object
   * @param {Array<string>|Object} propertiesToRemove - List of properties to remove
   */
  static removeStyleProperties(rPr, pPr, bodyPr, propertiesToRemove) {
    const props = Array.isArray(propertiesToRemove)
      ? propertiesToRemove
      : typeof propertiesToRemove === 'object' && propertiesToRemove !== null
        ? Object.keys(propertiesToRemove)
        : [propertiesToRemove]

    for (const prop of props) {
      switch (prop) {
        case 'fontSize':
          if (rPr) delete rPr['@_sz']
          break
        case 'color':
          if (rPr) delete rPr['a:solidFill']
          break
        case 'highlight':
          if (rPr) delete rPr['a:highlight']
          break
        case 'underline':
          if (rPr) delete rPr['@_u']
          break
        case 'bold':
          if (rPr) delete rPr['@_b']
          break
        case 'italic':
          if (rPr) delete rPr['@_i']
          break
        case 'strikethrough':
        case 'strike':
          if (rPr) delete rPr['@_strike']
          break
        case 'superscript':
        case 'subscript':
          if (rPr) delete rPr['@_baseline']
          break
        case 'smallCaps':
          if (rPr && rPr['@_cap'] === 'small') delete rPr['@_cap']
          break
        case 'uppercase':
          if (rPr && rPr['@_cap'] === 'all') delete rPr['@_cap']
          break
        case 'charSpacing':
          if (rPr) delete rPr['@_spc']
          break
        case 'shadowColor':
        case 'textShadow':
        case 'blur':
        case 'shadowRadius':
        case 'shadowOffsetX':
        case 'shadowOffsetY':
          if (rPr) delete rPr['a:effectLst']
          break
        case 'textAlignment':
        case 'alignment':
          if (pPr) delete pPr['@_algn']
          break
        case 'indentLevel':
        case 'bulletLevel':
          if (pPr) delete pPr['@_lvl']
          break
        case 'lineSpacing':
          if (pPr) delete pPr['a:lnSpc']
          break
        case 'bulletChar':
          if (pPr) delete pPr['a:buChar']
          break
        case 'bulletColor':
          if (pPr) delete pPr['a:buClr']
          break
        case 'bulletFontSize':
          if (pPr) {
            delete pPr['a:buSzPts']
            delete pPr['a:buSzPct']
          }
          break
        case 'bulletStyle':
          if (pPr) {
            delete pPr['a:buNone']
            delete pPr['a:buAutoNum']
            delete pPr['a:buChar']
          }
          break
        case 'verticalAlignment':
          if (bodyPr) delete bodyPr['@_anchor']
          break
        case 'textDirection':
          if (bodyPr) delete bodyPr['@_vert']
          break
        default:
          break
      }
    }
  }
}

module.exports = {
  StyleBuilder,
  normalizeHexColor,
  applyTextCasing,
}
