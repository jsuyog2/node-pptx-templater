---
name: pptx-engine-unified
description: Master skill for node-pptx-templater v3.0.0-alpha.0. Covers legacy code refactoring, AST OpenXML text reconstruction, spatial unit coordinate handling, dynamic table pagination, and multi-target export adapters.
version: 3.0.0-alpha.0
---

# PPTX Engine Unified Skill (v3.0.0-alpha.0)

## Mission
Guide the development and modernization of `node-pptx-templater` v3.0.0-alpha.0. This skill ensures strict OpenXML schema compliance, eliminates legacy synchronous blocking operations, and establishes a spatial layout and pagination engine.

---

## Technical Directives

### 1. I/O & Memory Pipeline
- Never use `fs.readFileSync` or `fs.writeFileSync`.
- Run all archive manipulations inside an in-memory Virtual File System (`Map<string, Uint8Array>`).
- Accept input from PPTX files, ZIP archives, and uncompressed OpenXML directory structures.
- Support 6 export targets: File, Stream, Buffer, ArrayBuffer, Base64, and Express HTTP responses.

### 2. AST-Based Run Reconstruction
- Do not use regexes to swap placeholders inside raw XML strings.
- Traverse paragraph nodes (`<a:p>`), combine fragmented child text runs (`<a:r>`), evaluate expressions, and reconstruct valid OpenXML runs with preserved typography.

### 3. Spatial Math & Unit Standards
- Store all coordinates and dimensions internally as integer EMUs:
  $$\text{EMU} = \text{Points} \times 12,700 = \text{Inches} \times 914,400 = \text{Pixels (96 DPI)} \times 9,525$$
- Expose methods: `getBounds()`, `setBounds()`, `move()`, `resize()`, `bringToFront()`, and `sendToBack()`.

### 4. Table Engine & Slide Overflow Splitting
- Compute dynamic row heights based on font size, line count, cell padding, and column width.
- Check row limits against $Y_{\text{limit}} = \text{SlideHeight} - \text{BottomMargin}$.
- Split overflowing tables, deep-clone the host slide with clean relationship mappings, repeat header rows, and distribute rows across subsequent slides.

### 5. Chart & Embedded Excel Synchronization
- When editing `<c:chartSpace>` series and categories, update the corresponding embedded `.xlsx` file in `ppt/embeddings/` to prevent presentation corruption prompts.