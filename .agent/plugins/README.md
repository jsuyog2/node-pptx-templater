# Plugins Architecture & Extension System

> **Directory**: `.agent/plugins/`  
> **Purpose**: Extension mechanisms, custom manager hooks, and agent plugin specifications.

---

## 1. Engine Plugin Protocol

To introduce a new custom element manager or serialization plugin without modifying core engine files:

```javascript
class CustomElementManager {
  #xmlParser;
  constructor(xmlParser, options = {}) { this.#xmlParser = xmlParser; }
  process(slideObj, context) { /* AST inspection / mutation */ }
  flush(zipManager, vfs) { /* Commit parts to VFS prior to serialization */ }
}
```

### Best Practices:
1. **Never Bypass VFS**: Read and write exclusively through `VirtualFileSystem`.
2. **Preserve Relationships**: Register new parts in `[Content_Types].xml` and `_rels/`.
3. **Pure JavaScript**: Zero native dependencies or external binaries.
