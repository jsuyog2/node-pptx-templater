export const CATEGORIES = [
  'Import/Export',
  'Slides',
  'Text & Placeholders',
  'Tables',
  'Charts',
  'Shapes',
  'Images & Media',
  'Elements',
];

export const CATEGORY_DESCRIPTIONS = {
  'Import/Export':
    'Manage presentation ingest, serialization, in-memory VFS lifecycle, buffer and stream exports, and Flat OPC XML conversions without disk I/O.',
  'Slides':
    'Slide lifecycle management, addition, cloning, layout assignments, reordering, deletion, and slide master relationship bindings.',
  'Text & Placeholders':
    'AST run-fragment healing, placeholder text substitution, inline text mutations, character formatting, and bullet styling.',
  'Tables':
    'Dynamic table generation, auto-pagination EMU calculation, repeating headers, cell styling, and row/column mutations.',
  'Charts':
    'Update DrawingML charts with simultaneous embedded Excel spreadsheet (XLSX) dual-sync to prevent Microsoft PowerPoint repair warnings.',
  'Shapes':
    'Insert geometric shapes, customize presets, fills, strokes, coordinates, and DrawingML transformations in EMU units.',
  'Images & Media':
    'Inject and replace high-resolution raster images, SVGs, and audio/video media with automatic Content Types and relationship mappings.',
  'Elements':
    'Low-level PresentationML shape tree (<p:spTree>) traversal, Z-order reindexing, spatial geometry math, and XML node inspection.',
};

export const CATEGORY_SLUG_MAP = {
  'Import/Export': 'import-export',
  'Slides': 'slides',
  'Text & Placeholders': 'text-placeholders',
  'Tables': 'tables',
  'Charts': 'charts',
  'Shapes': 'shapes',
  'Images & Media': 'images-media',
  'Elements': 'elements',
};

export const SLUG_TO_CATEGORY_MAP = {
  'import-export': 'Import/Export',
  'slides': 'Slides',
  'text-placeholders': 'Text & Placeholders',
  'tables': 'Tables',
  'charts': 'Charts',
  'shapes': 'Shapes',
  'images-media': 'Images & Media',
  'elements': 'Elements',
  // Aliases
  'lifecycle': 'Import/Export',
  'text': 'Text & Placeholders',
  'images': 'Images & Media',
};
