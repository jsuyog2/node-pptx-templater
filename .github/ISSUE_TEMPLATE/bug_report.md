---
name: "🐛 Bug Report"
about: Report a bug or issue with slide parsing, table cell merging, or chart updating.
title: "[BUG] "
labels: bug
assignees: ""
---

### 🚨 Description

A clear and concise description of the bug. Mention what slide features or XML operations were being run.

### 💻 Code Example

Provide a minimal code snippet to reproduce the behavior:

```js
const { PPTXTemplater } = require('node-pptx-templater');

// Code that triggers the error
```

### 📂 Presentation Details

- **Slide Count**:
- **Contains Charts?** (Yes/No)
- **Contains Tables?** (Yes/No)
- **Placeholders format**: (e.g. `{{name}}`)

### 📋 Error Log & Diagnostics

Paste your terminal error logs or safeParseXml diagnostic output here:

```text
// Paste diagnostic error object
```

### ⚙️ Environment Details

- **Library Version**:
- **Node.js Version**:
- **OS Version**: (e.g. Windows, macOS, Linux)
