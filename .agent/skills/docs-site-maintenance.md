# Skill: Documentation Portal Maintenance & Schema Synchronization

> **Domain**: Next.js App Router Static Site Generation & Cloudflare Pages Deployment  
> **Target Subsystems**: `docs-site/`, `scripts/generate-next-data.js`, `tests/unit/DocCoverage.test.js`

---

## 1. Documentation Architecture Overview

The documentation site at `docs-site/` is a high-performance Next.js 14+ static export deployed to Cloudflare Pages (`https://node-pptx-templater.pages.dev`).

All documentation content is completely decoupled from UI code and driven by JSON data schemas located in `docs-site/src/data/`:

| Schema File | Purpose | Structure / Key Fields |
| :--- | :--- | :--- |
| `apiReference.json` | Comprehensive catalog of all 103 public engine methods | `version`, `totalMethods`, `categories`, `methodsByCategory[category][]` |
| `features.json` | Granular feature inventory and capability matrix | `version`, `featureGroups[]`, `capabilities[]` |
| `docsNavigation.json`| Hierarchy for the fixed non-scrolling documentation sidebar | `sections[]`, `title`, `href`, `items[]` |
| `compatibility.json` | Runtime compatibility matrices (Node, Workers, Lambda, Deno) | `runtimes[]`, `supportedVersions`, `testedStatus` |
| `siteMeta.json` | Metadata, OpenGraph cards, title tags, and repository links | `title`, `description`, `domain`, `githubUrl` |

---

## 2. Playbook: Updating Documentation When Adding API Methods

Whenever a method or parameter is added, modified, or deprecated in `PPTXTemplater.js`:

### Step 1: Ensure JSDoc Completeness in `src/core/PPTXTemplater.js`
The automated doc coverage validator requires comprehensive JSDoc on every public method:
```javascript
/**
 * Short description of the method.
 *
 * @param {Type} paramName - Parameter description.
 * @returns {PPTXTemplater} this (chainable)
 *
 * @example
 * ppt.useSlide(1).myNewMethod();
 */
```

### Step 2: Update `docs-site/src/data/apiReference.json`
Add or update the method object under the appropriate category:
```json
{
  "id": "mynewmethod",
  "name": "myNewMethod",
  "signature": "ppt.myNewMethod(options = {})",
  "category": "Slides",
  "description": "Clear explanation of what the method does.",
  "openXmlImpact": "Details on which OpenXML parts are read or mutated.",
  "edgeCases": "Specific constraints, error modes, or repair edge cases.",
  "isStatic": false,
  "isAsync": true,
  "parameters": [
    {
      "name": "options",
      "type": "Object",
      "required": false,
      "default": "{}",
      "description": "Configuration object."
    }
  ],
  "returns": {
    "type": "Promise<PPTXTemplater>",
    "description": "Returns this for fluent chaining."
  },
  "exampleCode": "const ppt = await PPTXTemplater.load('deck.pptx');\nawait ppt.useSlide(1).myNewMethod();"
}
```

### Step 3: Run Automated Verification Tests
Verify that all public methods are documented:
```bash
npm run docs:validate
```
This runs `tests/unit/DocCoverage.test.js` to ensure zero undocumented methods, valid `sitemap.xml`, and valid `robots.txt`.

### Step 4: Run Data Generators (If Applicable)
If updating aggregated stats or documentation indices:
```bash
npm run docs:data
```

### Step 5: Build and Verify Static Export
Build the Next.js static output:
```bash
npm run docs:build
```
Ensure that `docs-site/out/` is generated without errors.

Test the static build locally:
```bash
npm run docs:preview
```

---

## 3. UI System & Style Guardrails

1. **Theme Support**:
   - UI supports Light and Dark modes via `ThemeContext.jsx` and `ThemeToggle.jsx`.
   - Use Tailwind CSS `dark:` variant classes across all components.
2. **Glassmorphism Aesthetic**:
   - Modern semi-transparent glass cards: `backdrop-blur-md bg-white/70 dark:bg-slate-900/70 border border-slate-200/50 dark:border-slate-800/50`.
3. **Animations**:
   - Smooth layout and card transitions via GSAP (`src/animations/gsapTransitions.js`).
4. **Layout Invariant**:
   - Sticky Header (`Home`, `Docs`, `Features`).
   - Fixed Non-Scrolling Sidebar on `/docs` routes (`Overview`, `Installation Guides`, `Compatibility`, `API Reference`).
