# OpenXML, ZIP, and Presentation Integrity

PPTX is a ZIP package containing interlinked OpenXML parts. A successful method call is not proof that the resulting file opens correctly. Every change must preserve package consistency.

## Package invariants
- Keep `[Content_Types].xml` overrides/defaults synchronized with package parts.
- Keep root, presentation, slide, layout, master, chart, media, workbook, and other `.rels` files consistent with the parts they reference.
- When cloning/importing slides or elements, remap relationship IDs and target paths to prevent collisions and dangling references.
- Keep `ppt/presentation.xml` slide order and relationship references synchronized after slide insertion, deletion, duplication, import, or reordering.
- Do not reuse non-visual shape IDs or other identifiers when uniqueness is required.
- Preserve themes, layouts, masters, notes, embedded media, chart parts, and workbook references unless the operation explicitly changes them.
- Remove relationships or package parts only when they are truly orphaned and safe to remove; account for shared media and relationships.
- Preserve unrelated XML nodes and namespaces. Avoid lossy parse/build cycles and broad serialization changes.
- Escape text and attribute values correctly; preserve Unicode and XML-special characters.
- Respect the distinction between presentation coordinates and units. Centralize conversions and verify the project's established conventions before changing them.
- For table cloning/insertion, preserve required OpenXML metadata such as unique row IDs where applicable and supported by the implementation.

## Relationship-aware operations
For slide duplication, slide import, image replacement, chart updates, hyperlink changes, and table operations:
1. Identify every affected part and relationship.
2. Update all relevant XML and relationship files together.
3. Ensure new targets exist and deleted targets are no longer referenced.
4. Add regression coverage for relationship collisions, missing targets, repeated operations, and shared assets where applicable.

## Validation
- Prefer the project's validator and repair routines where applicable, but do not treat automatic repair as a substitute for fixing the source defect.
- Verify both structural validity and expected semantic output.
- Where feasible, unzip the generated PPTX and inspect the affected XML/relationships in tests.
- Test output after save/reload, not only in-memory state.
- For chart updates, verify chart XML and embedded workbook/cache synchronization when that feature is supported.
- For table merges, verify span/merge markers, cell counts, and PowerPoint row identifiers as relevant.
- For images, verify content type, media bytes, dimensions, and relationship targets.
- For slide operations, verify ordering, slide IDs, relationship mappings, content types, and preserved assets.
- Never claim Microsoft PowerPoint compatibility has been proven unless the test environment actually verified it. Prefer wording such as “structural validation passed” when that is what was tested.

## Malformed or hostile input
- Treat uploaded PPTX/ZIP/XML content as untrusted.
- Apply or preserve ZIP-bomb, path traversal, entity-expansion, XML parser, oversized-input, and malformed-archive protections where relevant.
- Never extract archive paths without verifying that resolved destinations remain within the intended directory.
- Avoid enabling dangerous XML entity expansion or external entity resolution.
- Set sensible limits when the existing API provides a mechanism; keep defaults and error behavior documented.
