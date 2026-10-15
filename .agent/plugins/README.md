# Plugins Architecture & Extension System

> **Directory**: `.agent/plugins/`  
> **Purpose**: Extension mechanisms, custom manager hooks, and agent plugin specifications.

---

## 1. Overview

`node-pptx-templater` is built around a decoupled manager architecture. Core capabilities (slides, tables, charts, shapes, media, Z-order) are implemented as isolated sub-managers that interface with the `VirtualFileSystem` and `XMLParser`.

The `.agent/plugins/` directory provides architectural guidelines, blueprints, and extension points for:
1. Extending `PPTXTemplater` with domain-specific plugins (e.g., custom SVG converters, GIS shape renderers, CAD imports).
2. Authoring workspace plugins for autonomous agents (e.g., custom AST linters, presentation diff viewers).

---

## 2. Engine Plugin Protocol

To introduce a new custom element manager or serialization plugin without modifying core engine files:

### Step 1: Sub-Manager Interface Contract
A plugin manager must accept `(xmlParser, options)` in its constructor and implement standard lifecycle hooks:
```javascript
class CustomElementManager {
  #xmlParser;

  constructor(xmlParser, options = {}) {
    this.#xmlParser = xmlParser;
  }

  /**
   * Called during PPTXTemplater.useSlide() or AST traversal
   */
  process(slideObj, context) {
    // Inspect or mutate slide AST nodes
  }

  /**
   * Called during OutputWriter.flush() prior to ZIP generation
   */
  flush(zipManager, vfs) {
    // Commit pending binary or XML parts into the package
  }
}
```

### Step 2: Registering with the Engine
Plugins can be registered dynamically at runtime:
```javascript
const { PPTXTemplater } = require('node-pptx-templater');

// Future extension hook:
// PPTXTemplater.use(MyCustomPlugin);
```

---

## 3. Agent Tool Plugins

Agent plugins located in `.agent/plugins/` provide reusable utilities for AI agents working in this repository:
- **`openxml-linter`**: Fast static checker inspecting generated XML parts for unclosed tags or invalid DrawingML namespaces.
- **`vfs-inspector`**: Prints visual tree graphs of in-memory VFS parts.
- **`diff-viewer`**: Compares two `.pptx` presentations part-by-part, showing colorized XML diffs.

---

## 4. Best Practices for Plugin Authors

1. **Never Bypass the VFS**:
   Plugins must never attempt synchronous filesystem I/O (`fs.readFileSync`). All operations must read from and write to the provided `VirtualFileSystem` instance.
2. **Preserve Relationships & Content Types**:
   If a plugin adds a new part to the package (e.g., a custom XML part or audio file), it must register an `<Override>` in `[Content_Types].xml` and a corresponding relationship in `_rels`.
3. **Pure JavaScript**:
   All plugins must adhere to the pure JavaScript runtime mandate (zero native binaries).
