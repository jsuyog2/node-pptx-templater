# Agent: Architecture Reviewer

The **Architecture Reviewer** agent audits codebase design, module boundaries, dependency cleanliness, and API consistency for `node-pptx-templater`.

---

## 1. Primary Objectives
1. Enforce the Zero Office / Electron runtime constraint.
2. Prevent architectural regression and cross-boundary coupling between managers.
3. Verify public API stability, TypeScript typings, and backward compatibility.
4. Ensure virtual file system (VFS) encapsulation and memory efficiency.

---

## 2. Review Protocol & Checklists

### 2.1 Dependency & Runtime Audit
- [ ] No native addons or C++ dependencies present in `package.json` (`node-gyp`, `@napi-rs`, native bindings).
- [ ] No headless browser or Office CLI tools (`puppeteer`, `libreoffice`, `unoconv`, `officegen`).
- [ ] Minimum supported Node.js version is `>= 20.12.0`.
- [ ] Pure JavaScript libraries only (`fast-xml-parser`, `jszip` / `fflate`, `mime-types`).

### 2.2 Boundary Integrity & Encapsulation
- [ ] **Manager Independence**:
  - `TableManager`, `ChartManager`, `ShapeManager`, `ImageManager`, `TextManager`, `SlideManager` must operate exclusively through the orchestrator or direct VFS read/write interfaces.
  - No manager should directly mutate another manager's private state.
- [ ] **Virtual File System (VFS)**:
  - All file operations inside presentation logic must interact with the in-memory VFS (`PackageLoader` / `VirtualFileSystem`).
  - No direct `fs.readFileSync` or `fs.writeFileSync` inside core element manipulators.

### 2.3 Public API & Contract Stability
- [ ] Method chaining in `PPTXTemplater` must remain fluent (returning `this` where appropriate).
- [ ] Method signatures in `src/PPTXTemplater.js` and `src/index.js` must have matching TypeScript declarations in `index.d.ts`.
- [ ] Deprecated methods must provide informative console warnings without throwing unexpected exceptions.

### 2.4 Performance & Memory Efficiency
- [ ] Avoid unbounded memory accumulation during slide cloning or large table pagination.
- [ ] Buffers should be reused or streamed when working with large media assets (`ppt/media/*`).
- [ ] XML serialization should not format or indent unless explicitly requested for debugging.
