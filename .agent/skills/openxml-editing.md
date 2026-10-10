# Skill: OpenXML AST Editing & Serialization

This skill provides step-by-step instructions for safely inspecting, mutating, and serializing OpenXML DOM/AST structures within `node-pptx-templater`.

---

## 1. Fast-XML-Parser Configuration Standards

The library uses `fast-xml-parser` for bidirectional XML parsing and building. Always configure parsers and builders with these mandatory options:

```javascript
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

export const PARSER_OPTIONS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  allowBooleanAttributes: true,
  preserveOrder: false,
  parseTagValue: false, // Prevents automatic numeric conversion of phone numbers, row IDs, or codes
  trimValues: false,     // Critical to preserve whitespace inside <a:t>
  cdataPropName: '__cdata',
  textNodeName: '#text',
  ignoreDeclaration: false
};

export const BUILDER_OPTIONS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  cdataPropName: '__cdata',
  format: false,         // Keep XML compact; format: true can inject unwanted whitespace into text runs
  suppressEmptyNode: false
};
```

---

## 2. Navigating the AST: Handling Single Objects vs. Arrays

`fast-xml-parser` maps elements with one occurrence as a single Object, and elements with multiple occurrences as an Array.

### Best Practice: The Array Normalizer Helper
Always normalize child collections before iteration:

```javascript
export function ensureArray(element) {
  if (!element) return [];
  return Array.isArray(element) ? element : [element];
}

// Example: Iterating paragraphs inside a shape
const txBody = shape['p:txBody'];
if (txBody && txBody['a:p']) {
  const paragraphs = ensureArray(txBody['a:p']);
  for (const p of paragraphs) {
    const runs = ensureArray(p['a:r']);
    // Process runs safely...
  }
}
```

---

## 3. Preserving Namespaces & Attributes

OpenXML relies strictly on namespace prefixes:
- `p:` — PresentationML (`http://schemas.openxmlformats.org/presentationml/2006/main`)
- `a:` — DrawingML Main (`http://schemas.openxmlformats.org/drawingml/2006/main`)
- `r:` — Relationships (`http://schemas.openxmlformats.org/officeDocument/2006/relationships`)
- `c:` — Chart DrawingML (`http://schemas.openxmlformats.org/drawingml/2006/chart`)
- `a16:` — DrawingML 2014 Extension (`http://schemas.microsoft.com/office/drawing/2014/main`)

### Rules:
1. **Never strip or rename prefixes**: Always preserve the `p:`, `a:`, `r:`, or `c:` prefix on both element names and attributes (e.g., `r:id`).
2. **Attribute prefixes**: Attributes mapped by `fast-xml-parser` start with `@_` (e.g., `@_name`, `@_id`, `@_r:id`). When setting or mutating attributes:
   ```javascript
   node['@_name'] = 'UpdatedShape';
   node['@_r:id'] = newRelationshipId;
   ```
3. **Preserve Root Namespace Declarations**: Do not remove `xmlns:*` declarations on root elements (`<p:sld>`, `<p:presentation>`, `<c:chartSpace>`).

---

## 4. Spatial Coordinate Conversions (EMU)

Coordinates (`x`, `y`) and dimensions (`cx`, `cy`) in `<a:xfrm>` are strictly represented in EMUs (English Metric Units).

### Conversion Multipliers
```javascript
export const EMU_PER_POINT = 12700;
export const EMU_PER_INCH  = 914400;
export const EMU_PER_PIXEL = 9525; // at 96 DPI
export const EMU_PER_CM    = 360000;

export function pointsToEmu(pt) {
  return Math.round(pt * EMU_PER_POINT);
}

export function emuToPoints(emu) {
  return emu / EMU_PER_POINT;
}

export function inchesToEmu(inch) {
  return Math.round(inch * EMU_PER_INCH);
}

export function emuToInches(emu) {
  return emu / EMU_PER_INCH;
}

export function pixelsToEmu(px) {
  return Math.round(px * EMU_PER_PIXEL);
}

export function emuToPixels(emu) {
  return emu / EMU_PER_PIXEL;
}
```

**CRITICAL**: Always apply `Math.round()` when writing values back into `<a:off @_x="..." @_y="..."/>` or `<a:ext @_cx="..." @_cy="..."/>`. Floating-point values violate the OpenXML XSD and cause presentation corruption errors.

---

## 5. Modifying Text in `<a:t>`

PowerPoint text runs reside in `<a:r>` -> `<a:t>`.

### Whitespace Preservation
If replacement text contains leading or trailing spaces, OpenXML requires `xml:space="preserve"` on the `<a:t>` element:
```javascript
const textRun = {
  'a:rPr': { ...originalStyle },
  'a:t': {
    '#text': '  Padded Text  ',
    '@_xml:space': 'preserve'
  }
};
```
If there are no leading/trailing spaces, plain string assignment is permitted:
```javascript
run['a:t'] = 'Clean Text';
```

---

## 6. Verification Checklist for AST Edits

Before serializing the modified XML back into the VFS:
- [ ] No `undefined` or `null` values left in attributes or text nodes.
- [ ] All coordinates are rounded integer EMUs.
- [ ] Any added `r:id` references exist in the corresponding `.rels` file.
- [ ] Text inside tags has undergone run-healing if placeholders were targeted.
- [ ] Root namespaces are intact.
