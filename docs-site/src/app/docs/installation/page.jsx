import React from 'react';
import Link from 'next/link';
import { ArrowRight, CheckCircle2, Download, Terminal, Cpu, ShieldCheck } from 'lucide-react';
import CodeBlock from '../../../components/CodeBlock';

export const metadata = {
  title: 'Installation & Quickstart Guides',
  description: 'How to install and set up node-pptx-templater in Node.js, TypeScript, and serverless projects.',
};

const ESM_EXAMPLE = `// Using ECMAScript Modules (ESM)
import { PPTXTemplater } from 'node-pptx-templater';

const ppt = await PPTXTemplater.load('./template.pptx');
ppt.useSlide(1).replaceText({ '{{USER}}': 'Sarah Connor' });
await ppt.saveToFile('./output.pptx');`;

const CJS_EXAMPLE = `// Using CommonJS (CJS)
const { PPTXTemplater } = require('node-pptx-templater');

async function run() {
  const ppt = await PPTXTemplater.load('./template.pptx');
  ppt.useSlide(1).replaceText({ '{{USER}}': 'Sarah Connor' });
  await ppt.saveToFile('./output.pptx');
}

run();`;

const TS_EXAMPLE = `// TypeScript with native autocompletion and types
import { PPTXTemplater, PPTXSlide, ReplaceTextOptions } from 'node-pptx-templater';

async function generateDeck(): Promise<Buffer> {
  const ppt = await PPTXTemplater.load('./deck.pptx');
  const slide: PPTXSlide = ppt.useSlide(1);
  
  slide.replaceText({ '{{METRIC}}': '99.98%' }, { caseSensitive: true });
  return await ppt.toBuffer();
}`;

export default function InstallationPage() {
  return (
    <article className="space-y-12">
      {/* Page Title */}
      <div>
        <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full glass-pill text-xs font-semibold mb-4">
          <Download className="w-3.5 h-3.5" />
          <span>Installation Guide</span>
        </div>
        <h1 className="text-3xl sm:text-4xl font-extrabold tracking-tight text-slate-900 dark:text-white">
          Installation &amp; Project Setup
        </h1>
        <p className="mt-3 text-base sm:text-lg text-slate-600 dark:text-slate-300 leading-relaxed">
          Install `node-pptx-templater` via your favorite package manager. Ships with dual ESM/CJS builds and native TypeScript declarations.
        </p>
      </div>

      {/* Package Manager Commands */}
      <section className="space-y-4">
        <h2 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900 dark:text-white flex items-center gap-2">
          <Terminal className="w-5 h-5 text-brand-500" />
          <span>Package Installation</span>
        </h2>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="glass-card p-4 rounded-xl space-y-2">
            <span className="text-xs font-semibold text-slate-500">npm</span>
            <CodeBlock code="npm install node-pptx-templater" language="bash" />
          </div>

          <div className="glass-card p-4 rounded-xl space-y-2">
            <span className="text-xs font-semibold text-slate-500">pnpm</span>
            <CodeBlock code="pnpm add node-pptx-templater" language="bash" />
          </div>

          <div className="glass-card p-4 rounded-xl space-y-2">
            <span className="text-xs font-semibold text-slate-500">yarn</span>
            <CodeBlock code="yarn add node-pptx-templater" language="bash" />
          </div>

          <div className="glass-card p-4 rounded-xl space-y-2">
            <span className="text-xs font-semibold text-slate-500">bun</span>
            <CodeBlock code="bun add node-pptx-templater" language="bash" />
          </div>
        </div>
      </section>

      {/* Prerequisites */}
      <section className="space-y-4">
        <h2 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900 dark:text-white">
          Prerequisites &amp; System Requirements
        </h2>

        <div className="glass-card p-6 rounded-2xl space-y-4">
          <div className="flex items-start gap-3">
            <CheckCircle2 className="w-5 h-5 text-emerald-500 shrink-0 mt-0.5" />
            <div>
              <h4 className="text-sm font-bold text-slate-900 dark:text-white">
                Node.js Version: &gt;= 18.0.0 (Recommended: &gt;= 20.12.0)
              </h4>
              <p className="text-xs text-slate-600 dark:text-slate-400 mt-1">
                Runs on modern V8 JavaScript runtimes. Fully compatible with Node 18, 20, 22, 24, AWS Lambda, Cloudflare Workers, and Bun.
              </p>
            </div>
          </div>

          <div className="flex items-start gap-3">
            <CheckCircle2 className="w-5 h-5 text-emerald-500 shrink-0 mt-0.5" />
            <div>
              <h4 className="text-sm font-bold text-slate-900 dark:text-white">
                Zero Microsoft Office or Native Addons Required
              </h4>
              <p className="text-xs text-slate-600 dark:text-slate-400 mt-1">
                No `node-gyp`, no Python build toolchains, no LibreOffice, and no Java JDK required. Pure JavaScript AST and ZIP parsing.
              </p>
            </div>
          </div>
        </div>
      </section>

      {/* Module Imports (ESM vs CJS) */}
      <section className="space-y-6">
        <h2 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900 dark:text-white">
          Module Configuration
        </h2>

        <div className="space-y-4">
          <h3 className="text-sm font-bold text-slate-900 dark:text-white">1. Modern ECMAScript Modules (ESM)</h3>
          <CodeBlock code={ESM_EXAMPLE} language="javascript" filename="index.mjs" />
        </div>

        <div className="space-y-4">
          <h3 className="text-sm font-bold text-slate-900 dark:text-white">2. CommonJS (CJS)</h3>
          <CodeBlock code={CJS_EXAMPLE} language="javascript" filename="index.js" />
        </div>

        <div className="space-y-4">
          <h3 className="text-sm font-bold text-slate-900 dark:text-white">3. TypeScript Support</h3>
          <p className="text-xs text-slate-600 dark:text-slate-400">
            TypeScript definitions (`index.d.ts`) are bundled directly with the package. No `@types/node-pptx-templater` installation is needed.
          </p>
          <CodeBlock code={TS_EXAMPLE} language="typescript" filename="report.ts" />
        </div>
      </section>

      {/* Next Step CTA */}
      <div className="pt-8 border-t border-slate-200 dark:border-slate-800 flex flex-wrap items-center justify-between gap-4">
        <div>
          <span className="text-xs text-slate-500">Next Step</span>
          <p className="text-sm font-semibold text-slate-900 dark:text-white">Review Runtime Compatibility &amp; Serverless</p>
        </div>
        <Link
          href="/docs/compatibility"
          className="flex items-center gap-2 px-5 py-2.5 rounded-xl text-xs font-semibold bg-brand-600 text-white hover:bg-brand-500 shadow-md shadow-brand-500/20 transition-all"
        >
          <span>Compatibility Matrix</span>
          <ArrowRight className="w-3.5 h-3.5" />
        </Link>
      </div>
    </article>
  );
}
