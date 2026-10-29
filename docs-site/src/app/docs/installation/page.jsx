'use client';

import React from 'react';
import Link from 'next/link';
import { ArrowRight, Terminal, CheckCircle2, Download, Cpu, ShieldCheck } from 'lucide-react';
import PowerPointSlideDeck from '../../../components/PowerPointSlideDeck';
import CodeBlock from '../../../components/CodeBlock';

const INSTALLATION_SLIDES = [
  {
    id: 'install-1',
    category: 'Installation',
    title: 'Package Manager Commands',
    subtitle: 'Zero Native Compilation • Ships Pre-Bundled',
    signature: 'npm / pnpm / yarn / bun',
    content: (
      <div className="space-y-4">
        <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)]">
          Install <code className="font-mono text-[var(--ppt-text-primary)]">node-pptx-templater</code> via your preferred package manager. No node-gyp or C/C++ toolchains needed:
        </p>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)]">
            <span className="text-xs font-bold text-[#D24726]">npm</span>
            <CodeBlock code="npm install node-pptx-templater" language="bash" />
          </div>

          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)]">
            <span className="text-xs font-bold text-[#D24726]">pnpm</span>
            <CodeBlock code="pnpm add node-pptx-templater" language="bash" />
          </div>

          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)]">
            <span className="text-xs font-bold text-[#D24726]">yarn</span>
            <CodeBlock code="yarn add node-pptx-templater" language="bash" />
          </div>

          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)]">
            <span className="text-xs font-bold text-[#D24726]">bun</span>
            <CodeBlock code="bun add node-pptx-templater" language="bash" />
          </div>
        </div>
      </div>
    ),
  },
  {
    id: 'install-2',
    category: 'Modules',
    title: 'Dual ESM and CommonJS Support',
    subtitle: 'First-Class Compatibility Across Modern and Legacy Node.js',
    signature: 'Dual Build Exports',
    content: (
      <div className="space-y-3">
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
          <div>
            <div className="text-xs font-bold text-[var(--ppt-text-primary)] mb-1">
              ECMAScript Modules (ESM)
            </div>
            <CodeBlock
              code={`import { PPTXTemplater } from 'node-pptx-templater';

const ppt = await PPTXTemplater.load('./deck.pptx');
ppt.useSlide(1).replaceText({ '{{USER}}': 'Sarah' });
await ppt.saveToFile('./output.pptx');`}
              language="javascript"
              filename="index.mjs"
            />
          </div>

          <div>
            <div className="text-xs font-bold text-[var(--ppt-text-primary)] mb-1">
              CommonJS (CJS)
            </div>
            <CodeBlock
              code={`const { PPTXTemplater } = require('node-pptx-templater');

async function run() {
  const ppt = await PPTXTemplater.load('./deck.pptx');
  ppt.useSlide(1).replaceText({ '{{USER}}': 'Sarah' });
  await ppt.saveToFile('./output.pptx');
}
run();`}
              language="javascript"
              filename="index.cjs"
            />
          </div>
        </div>
      </div>
    ),
  },
  {
    id: 'install-3',
    category: 'TypeScript',
    title: 'Native TypeScript Declarations',
    subtitle: 'Strong Typings & IDE Autocompletion Out of the Box',
    signature: 'index.d.ts',
    content: (
      <div className="space-y-3">
        <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)]">
          TypeScript definitions are bundled inside the npm package under <code className="font-mono text-[#D24726]">index.d.ts</code> with JSDoc descriptions for every method:
        </p>
        <CodeBlock
          code={`import { PPTXTemplater, PPTXSlide, ReplaceTextOptions } from 'node-pptx-templater';

async function generateReport(): Promise<Buffer> {
  const ppt = await PPTXTemplater.load('./templates/executive.pptx');
  const slide: PPTXSlide = ppt.useSlide(1);

  const options: ReplaceTextOptions = { caseSensitive: true };
  slide.replaceText({ '{{METRIC}}': '99.98%' }, options);

  return await ppt.toBuffer();
}`}
          language="typescript"
          filename="report-generator.ts"
        />
      </div>
    ),
  },
];

export default function InstallationPage() {
  return (
    <PowerPointSlideDeck
      slides={INSTALLATION_SLIDES}
      categoryTitle="Installation"
    />
  );
}
