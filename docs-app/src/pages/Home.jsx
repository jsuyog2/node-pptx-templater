import React, { useState } from 'react';
import { Terminal, Shield, Cpu, Zap, Layers, BarChart3, Table, FileSpreadsheet, ArrowRight, CheckCircle2, XCircle, Copy, Check } from 'lucide-react';
import CodeBlock from '../components/CodeBlock';
import SEOHead from '../components/SEOHead';

export default function Home({ onNavigate }) {
  const [activeTab, setActiveTab] = useState('cjs');
  const [copiedInstall, setCopiedInstall] = useState(false);

  const handleCopyInstall = () => {
    navigator.clipboard.writeText('npm install node-pptx-templater');
    setCopiedInstall(true);
    setTimeout(() => setCopiedInstall(false), 2000);
  };

  const cjsCode = `const { PPTXTemplater } = require('node-pptx-templater');

async function main() {
  // 1. Ingest template PPTX file or Buffer
  const ppt = await PPTXTemplater.load('./templates/executive-deck.pptx');

  // 2. Perform safe, run-fragment-healed text replacements
  ppt.useSlide(1).replaceText({
    '{{COMPANY_NAME}}': 'Acme Global Technologies Inc.',
    '{{REPORT_QUARTER}}': 'Q3 2026',
    '{{PRESENTER}}': 'Sarah Jenkins, VP Engineering'
  });

  // 3. Update native chart series & automatically sync embedded Excel workbook
  ppt.useSlide(2);
  await ppt.updateChart('ARRPerformanceChart', {
    categories: ['Q1', 'Q2', 'Q3', 'Q4 (Proj)'],
    series: [
      { name: 'Cloud ARR', values: [18.4, 22.1, 26.8, 31.5] },
      { name: 'AI Services', values: [8.2, 12.6, 17.4, 23.0] }
    ]
  });

  // 4. Save without Microsoft PowerPoint repair prompts
  await ppt.saveToFile('./output/Final_Quarterly_Review.pptx');
  console.log('✓ Presentation generated with zero XML corruption!');
}

main();`;

  const esmCode = `import { PPTXTemplater } from 'node-pptx-templater';

// Perfect for AWS Lambda, Cloudflare Workers, or Next.js Route Handlers
export async function generateDeck(payload) {
  const ppt = await PPTXTemplater.load(payload.templateBuffer);

  // Manipulate all slides concurrently
  ppt.useSlide('all').replaceText(payload.replacements);

  // Direct export to in-memory Buffer (no disk I/O required)
  const outputBuffer = await ppt.toBuffer();
  return outputBuffer;
}`;

  return (
    <div>
      <SEOHead
        title="node-pptx-templater — High-Performance PowerPoint OpenXML Template Engine"
        description="Populate PowerPoint templates dynamically in Node.js & Serverless runtimes without Microsoft Office dependencies."
        canonicalUrl="https://node-pptx-templater.pages.dev/"
      />

      {/* Hero Section */}
      <section className="relative pt-8 pb-16 sm:pt-14 sm:pb-24 overflow-hidden text-center sm:text-left">
        <div className="max-w-5xl mx-auto">
          {/* Release & Build Badges */}
          <div className="flex flex-wrap items-center justify-center sm:justify-start gap-2 mb-6">
            <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-blue-500/10 text-blue-600 dark:text-blue-400 border border-blue-500/20">
              <Zap className="w-3.5 h-3.5" /> v3.0.0-alpha.0 Architecture
            </span>
            <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border border-emerald-500/20">
              <Shield className="w-3.5 h-3.5" /> 285+ Tests Passing
            </span>
            <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-purple-500/10 text-purple-600 dark:text-purple-400 border border-purple-500/20">
              <Cpu className="w-3.5 h-3.5" /> Zero Native Binaries
            </span>
          </div>

          {/* Headline */}
          <h1 className="text-3xl sm:text-5xl lg:text-6xl font-extrabold tracking-tight text-slate-900 dark:text-slate-50 leading-[1.15] mb-6">
            High-Performance{' '}
            <span className="bg-gradient-to-r from-blue-600 via-indigo-600 to-purple-600 dark:from-blue-400 dark:via-indigo-300 dark:to-purple-400 bg-clip-text text-transparent">
              PowerPoint OpenXML
            </span>{' '}
            Engine for Node.js
          </h1>

          {/* Subtitle */}
          <p className="text-base sm:text-lg lg:text-xl text-slate-600 dark:text-slate-300 max-w-3xl leading-relaxed mb-8">
            Populate PowerPoint templates dynamically in Node.js and Serverless runtimes.
            Heals split run fragments, synchronizes native charts with embedded Excel workbooks,
            and merges table cells without Microsoft Office repair warnings.
          </p>

          {/* Install Bar & Primary Actions */}
          <div className="flex flex-col sm:flex-row items-center gap-3 sm:gap-4 max-w-xl">
            <div className="w-full sm:flex-1 flex items-center justify-between px-4 py-2.5 rounded-xl bg-slate-900 text-slate-100 border border-slate-800 font-mono text-xs sm:text-sm shadow-md">
              <div className="flex items-center gap-2 overflow-x-auto">
                <span className="text-slate-500 select-none">$</span>
                <span>npm install node-pptx-templater</span>
              </div>
              <button
                onClick={handleCopyInstall}
                aria-label="Copy install command"
                className="ml-2 p-1.5 rounded-md hover:bg-slate-800 text-slate-400 hover:text-white transition-colors"
              >
                {copiedInstall ? (
                  <Check className="w-4 h-4 text-emerald-400" />
                ) : (
                  <Copy className="w-4 h-4" />
                )}
              </button>
            </div>

            <button
              onClick={() => onNavigate('/api-reference')}
              className="w-full sm:w-auto px-6 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-500 text-white font-semibold text-xs sm:text-sm shadow-md hover:shadow-lg transition-all flex items-center justify-center gap-2 group whitespace-nowrap"
            >
              <span>Explore 102+ APIs</span>
              <ArrowRight className="w-4 h-4 group-hover:translate-x-0.5 transition-transform" />
            </button>
          </div>
        </div>
      </section>

      {/* Feature Comparison Matrix */}
      <section className="my-16">
        <div className="text-center sm:text-left mb-6">
          <h2 className="text-xl sm:text-2xl font-bold text-slate-900 dark:text-slate-100 tracking-tight">
            Why node-pptx-templater?
          </h2>
          <p className="text-xs sm:text-sm text-slate-500 dark:text-slate-400 mt-1">
            Compare architectural capabilities against legacy PowerPoint libraries.
          </p>
        </div>

        <div className="overflow-x-auto glass-card rounded-2xl border border-slate-200/80 dark:border-slate-800/80">
          <table className="w-full text-left text-xs sm:text-sm border-collapse">
            <thead>
              <tr className="bg-slate-100/70 dark:bg-slate-800/50 border-b border-slate-200 dark:border-slate-800 text-slate-700 dark:text-slate-300 font-semibold">
                <th className="py-3 px-4">Feature / Capability</th>
                <th className="py-3 px-4 text-blue-600 dark:text-blue-400 font-bold bg-blue-50/50 dark:bg-blue-950/20">
                  node-pptx-templater
                </th>
                <th className="py-3 px-4 text-slate-500">pptxgenjs</th>
                <th className="py-3 px-4 text-slate-500">officegen</th>
                <th className="py-3 px-4 text-slate-500">LibreOffice CLI</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200/60 dark:divide-slate-800/60 text-slate-600 dark:text-slate-400">
              <tr className="hover:bg-slate-50/50 dark:hover:bg-slate-800/20">
                <td className="py-3 px-4 font-medium text-slate-800 dark:text-slate-200">
                  Template-Based Editing (Preserves Visual Theme)
                </td>
                <td className="py-3 px-4 font-semibold text-emerald-600 dark:text-emerald-400 bg-blue-50/30 dark:bg-blue-950/10">
                  <CheckCircle2 className="w-4 h-4 inline mr-1 text-emerald-500" /> Full In-Place Mutation
                </td>
                <td className="py-3 px-4 text-rose-500">
                  <XCircle className="w-4 h-4 inline mr-1" /> Blank Only (No Existing PPTX)
                </td>
                <td className="py-3 px-4 text-rose-500">
                  <XCircle className="w-4 h-4 inline mr-1" /> Blank Only
                </td>
                <td className="py-3 px-4 text-slate-400">Macro scripting only</td>
              </tr>
              <tr className="hover:bg-slate-50/50 dark:hover:bg-slate-800/20">
                <td className="py-3 px-4 font-medium text-slate-800 dark:text-slate-200">
                  Run Fragment Healing (Anti-Corruption)
                </td>
                <td className="py-3 px-4 font-semibold text-emerald-600 dark:text-emerald-400 bg-blue-50/30 dark:bg-blue-950/10">
                  <CheckCircle2 className="w-4 h-4 inline mr-1 text-emerald-500" /> AST Normalizer
                </td>
                <td className="py-3 px-4 text-rose-500">N/A</td>
                <td className="py-3 px-4 text-rose-500">N/A</td>
                <td className="py-3 px-4 text-rose-500">Brittle Regex</td>
              </tr>
              <tr className="hover:bg-slate-50/50 dark:hover:bg-slate-800/20">
                <td className="py-3 px-4 font-medium text-slate-800 dark:text-slate-200">
                  Chart Excel Dual-Sync (.xlsx Caching)
                </td>
                <td className="py-3 px-4 font-semibold text-emerald-600 dark:text-emerald-400 bg-blue-50/30 dark:bg-blue-950/10">
                  <CheckCircle2 className="w-4 h-4 inline mr-1 text-emerald-500" /> XML + XLSX Sync
                </td>
                <td className="py-3 px-4 text-slate-400">Basic generated data</td>
                <td className="py-3 px-4 text-rose-500">Unsupported</td>
                <td className="py-3 px-4 text-slate-400">Slow headless recalc</td>
              </tr>
              <tr className="hover:bg-slate-50/50 dark:hover:bg-slate-800/20">
                <td className="py-3 px-4 font-medium text-slate-800 dark:text-slate-200">
                  Table Overflow Auto-Pagination
                </td>
                <td className="py-3 px-4 font-semibold text-emerald-600 dark:text-emerald-400 bg-blue-50/30 dark:bg-blue-950/10">
                  <CheckCircle2 className="w-4 h-4 inline mr-1 text-emerald-500" /> splitTableOnOverflow()
                </td>
                <td className="py-3 px-4 text-rose-500">Manual calculation</td>
                <td className="py-3 px-4 text-rose-500">Unsupported</td>
                <td className="py-3 px-4 text-rose-500">Unsupported</td>
              </tr>
              <tr className="hover:bg-slate-50/50 dark:hover:bg-slate-800/20">
                <td className="py-3 px-4 font-medium text-slate-800 dark:text-slate-200">
                  Serverless Ready (Lambda, Workers, Edge)
                </td>
                <td className="py-3 px-4 font-semibold text-emerald-600 dark:text-emerald-400 bg-blue-50/30 dark:bg-blue-950/10">
                  <CheckCircle2 className="w-4 h-4 inline mr-1 text-emerald-500" /> &lt;2MB Pure JS, &lt;15ms boot
                </td>
                <td className="py-3 px-4 text-emerald-600 dark:text-emerald-400">Pure JS</td>
                <td className="py-3 px-4 text-slate-400">High memory</td>
                <td className="py-3 px-4 text-rose-500">
                  <XCircle className="w-4 h-4 inline mr-1" /> Heavy 500MB+ Docker
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </section>

      {/* Interactive Quick Start Snippets */}
      <section className="my-16">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-4">
          <div>
            <h2 className="text-xl sm:text-2xl font-bold text-slate-900 dark:text-slate-100 tracking-tight">
              Interactive Quick Start
            </h2>
            <p className="text-xs sm:text-sm text-slate-500 dark:text-slate-400">
              Clean, fully verified examples ready to paste directly into your project.
            </p>
          </div>
          <div className="flex items-center p-1 rounded-xl bg-slate-100 dark:bg-slate-900 border border-slate-200 dark:border-slate-800 self-start sm:self-auto">
            <button
              onClick={() => setActiveTab('cjs')}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${
                activeTab === 'cjs'
                  ? 'bg-white dark:bg-slate-800 text-blue-600 dark:text-blue-400 shadow-sm'
                  : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-100'
              }`}
            >
              CommonJS
            </button>
            <button
              onClick={() => setActiveTab('esm')}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${
                activeTab === 'esm'
                  ? 'bg-white dark:bg-slate-800 text-blue-600 dark:text-blue-400 shadow-sm'
                  : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-100'
              }`}
            >
              ESM / Serverless
            </button>
          </div>
        </div>

        <CodeBlock
          code={activeTab === 'cjs' ? cjsCode : esmCode}
          language="javascript"
          title={activeTab === 'cjs' ? 'Node.js Script (CommonJS)' : 'Modern ESM / Serverless Handler'}
        />
      </section>

      {/* Core Architectural Pillars */}
      <section className="my-16">
        <div className="text-center sm:text-left mb-8">
          <h2 className="text-xl sm:text-2xl font-bold text-slate-900 dark:text-slate-100 tracking-tight">
            Core Architectural Superpowers
          </h2>
          <p className="text-xs sm:text-sm text-slate-500 dark:text-slate-400 mt-1">
            Engineered specifically to solve the hardest problems in OpenXML presentation generation.
          </p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {/* Card 1 */}
          <div className="glass-card rounded-2xl p-6 border border-slate-200/80 dark:border-slate-800/80">
            <div className="w-10 h-10 rounded-xl bg-blue-500/10 text-blue-600 dark:text-blue-400 flex items-center justify-center mb-4">
              <Zap className="w-5 h-5" />
            </div>
            <h3 className="text-base sm:text-lg font-bold text-slate-900 dark:text-slate-100 mb-2">
              Run Fragment Healing
            </h3>
            <p className="text-xs sm:text-sm text-slate-600 dark:text-slate-300 leading-relaxed mb-4">
              PowerPoint regularly fragments a single placeholder like <code>{'{{COMPANY}}'}</code> into three separate <code>&lt;a:r&gt;</code> run nodes. Our AST-level paragraph normalizer automatically detects and consolidates split runs without losing bold, color, or font attributes.
            </p>
            <button
              onClick={() => onNavigate('/guides')}
              className="text-xs font-semibold text-blue-600 dark:text-blue-400 hover:underline flex items-center gap-1"
            >
              Read Run Healing Guide <ArrowRight className="w-3 h-3" />
            </button>
          </div>

          {/* Card 2 */}
          <div className="glass-card rounded-2xl p-6 border border-slate-200/80 dark:border-slate-800/80">
            <div className="w-10 h-10 rounded-xl bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 flex items-center justify-center mb-4">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <h3 className="text-base sm:text-lg font-bold text-slate-900 dark:text-slate-100 mb-2">
              Excel Dual-Sync Technology
            </h3>
            <p className="text-xs sm:text-sm text-slate-600 dark:text-slate-300 leading-relaxed mb-4">
              Modifying only the XML chart cache causes PowerPoint to pop up an alarming "Repair Presentation" dialog when opened. node-pptx-templater mutates both the DrawingML chart cache and the embedded binary Excel workbook simultaneously.
            </p>
            <button
              onClick={() => onNavigate('/guides')}
              className="text-xs font-semibold text-blue-600 dark:text-blue-400 hover:underline flex items-center gap-1"
            >
              Read Excel Dual-Sync Guide <ArrowRight className="w-3 h-3" />
            </button>
          </div>

          {/* Card 3 */}
          <div className="glass-card rounded-2xl p-6 border border-slate-200/80 dark:border-slate-800/80">
            <div className="w-10 h-10 rounded-xl bg-purple-500/10 text-purple-600 dark:text-purple-400 flex items-center justify-center mb-4">
              <Table className="w-5 h-5" />
            </div>
            <h3 className="text-base sm:text-lg font-bold text-slate-900 dark:text-slate-100 mb-2">
              Auto-Pagination & Row Merging
            </h3>
            <p className="text-xs sm:text-sm text-slate-600 dark:text-slate-300 leading-relaxed mb-4">
              Tables with 50+ database rows automatically paginate across newly cloned slides with repeated header rows using <code>splitTableOnOverflow()</code>. Full support for vertical <code>rowspan</code> and horizontal <code>gridspan</code> merges.
            </p>
            <button
              onClick={() => onNavigate('/guides')}
              className="text-xs font-semibold text-blue-600 dark:text-blue-400 hover:underline flex items-center gap-1"
            >
              Read Table Guide <ArrowRight className="w-3 h-3" />
            </button>
          </div>

          {/* Card 4 */}
          <div className="glass-card rounded-2xl p-6 border border-slate-200/80 dark:border-slate-800/80">
            <div className="w-10 h-10 rounded-xl bg-amber-500/10 text-amber-600 dark:text-amber-400 flex items-center justify-center mb-4">
              <Layers className="w-5 h-5" />
            </div>
            <h3 className="text-base sm:text-lg font-bold text-slate-900 dark:text-slate-100 mb-2">
              Z-Order & Spatial Geometry
            </h3>
            <p className="text-xs sm:text-sm text-slate-600 dark:text-slate-300 leading-relaxed mb-4">
              Full control over visual stacking orders with <code>bringToFront</code>, <code>sendToBack</code>, and <code>stepForward</code>. Spatial unit converters convert English Metric Units (EMU), pixels, points, inches, and millimeters with zero rounding distortion.
            </p>
            <button
              onClick={() => onNavigate('/guides')}
              className="text-xs font-semibold text-blue-600 dark:text-blue-400 hover:underline flex items-center gap-1"
            >
              Read Spatial Engine Guide <ArrowRight className="w-3 h-3" />
            </button>
          </div>
        </div>
      </section>

      {/* CTA Section */}
      <section className="my-20 p-8 sm:p-12 rounded-3xl bg-gradient-to-tr from-blue-600 via-indigo-700 to-purple-800 text-white shadow-xl text-center">
        <h2 className="text-2xl sm:text-3xl font-extrabold tracking-tight mb-3">
          Ready to Automate PowerPoint Without Corruption?
        </h2>
        <p className="text-blue-100 text-xs sm:text-sm max-w-2xl mx-auto mb-8 leading-relaxed">
          Install the package or browse through our comprehensive catalog of 102 verified API methods.
        </p>
        <div className="flex flex-wrap items-center justify-center gap-3">
          <button
            onClick={() => onNavigate('/api-reference')}
            className="px-6 py-2.5 rounded-xl bg-white text-blue-700 hover:bg-blue-50 font-semibold text-xs sm:text-sm shadow transition-all"
          >
            Browse Full API Reference
          </button>
          <button
            onClick={() => onNavigate('/showcase')}
            className="px-6 py-2.5 rounded-xl bg-blue-500/30 hover:bg-blue-500/40 text-white border border-white/20 font-semibold text-xs sm:text-sm transition-all"
          >
            View 12 Showcase Demos
          </button>
        </div>
      </section>
    </div>
  );
}
