# Plugins Architecture for `node-pptx-templater`

The plugin system allows extending `node-pptx-templater` with custom element parsers, template interceptors, chart rendering adapters, and export pipeline extensions without modifying library core.

---

## 1. Plugin Interface Specification

A plugin is an object or class implementing a standard lifecycle interface:

```javascript
export interface PPTXPlugin {
  /** Unique plugin identifier */
  name: string;
  
  /** Semantic version string */
  version?: string;

  /** Hook called upon presentation loading into VFS */
  onLoad?(context: PluginLoadContext): Promise<void> | void;

  /** Hook called before slide template tokens are evaluated */
  beforeRenderSlide?(context: SlideRenderContext): Promise<void> | void;

  /** Hook called after slide elements are mutated */
  afterRenderSlide?(context: SlideRenderContext): Promise<void> | void;

  /** Hook called before presentation package is serialized */
  beforeSerialize?(context: SerializeContext): Promise<void> | void;

  /** Custom element processor */
  registerCustomElements?(registry: ElementRegistry): void;
}
```

---

## 2. Extension Points

### 2.1 Custom Element Parsers
Plugins can register handlers for custom OpenXML markup or vendor-specific namespaces:
- SmartArt manipulators
- Embedded 3D model processors (`<am3d:model>`)
- Ink drawing / signature processors

### 2.2 Template Interceptors
Interceptors can inspect and transform tokens dynamically before `TemplateEngine.js` performs run healing:
- Internationalization / localization filters (`{{t:greeting}}`)
- Format filters (`{{price | currency:'USD'}}`)
- Conditional logic blocks (`{{#if showChart}}...{{/if}}`)

### 2.3 Export Pipeline Adapters
Custom exporters can plug into the output pipeline:
- Direct S3 / Cloud Storage uploaders
- PDF conversion gateways (calling external rendering services)
- Webhook dispatchers upon presentation generation

---

## 3. Registering Plugins

Plugins are registered on a `PPTXTemplater` instance using `.use()`:

```javascript
import { PPTXTemplater } from 'node-pptx-templater';
import { CustomWatermarkPlugin } from './plugins/WatermarkPlugin.js';

const templater = new PPTXTemplater();
templater.use(new CustomWatermarkPlugin({ text: 'CONFIDENTIAL' }));

await templater.load('template.pptx');
await templater.process({ /* data */ });
await templater.save('output.pptx');
```

---

## 4. Developing New Plugins
1. Place experimental or internal plugins in `.agent/plugins/<plugin-name>/`.
2. Include comprehensive unit tests verifying that plugins do not violate OpenXML schema integrity or break VFS caching.
3. Export an idempotent factory function returning the plugin object.
