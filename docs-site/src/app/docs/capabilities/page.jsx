'use client';

import React from 'react';
import Link from 'next/link';
import { ArrowRight, Cpu, CheckCircle2, Zap, Server, Globe, ShieldCheck } from 'lucide-react';
import PowerPointSlideDeck from '../../../components/PowerPointSlideDeck';
import CodeBlock from '../../../components/CodeBlock';

const RUNTIME_MATRIX = [
  {
    name: 'Node.js (>= 18, 20.x, 22.x, 24.x)',
    tier: 'Tier 1 - Primary Support',
    status: 'Fully Supported',
    notes: 'CJS & ESM, native TS definitions, high-throughput in-memory VFS.',
    bootTime: '< 15 ms',
  },
  {
    name: 'AWS Lambda (Node.js 18 / 20 / 22)',
    tier: 'Tier 1 - Serverless',
    status: 'Fully Supported',
    notes: 'Zero /tmp disk writes. Stream directly to S3 or export Buffer.',
    bootTime: '< 25 ms',
  },
  {
    name: 'Cloudflare Workers (V8 Isolates)',
    tier: 'Tier 1 - Edge Computing',
    status: 'Fully Supported',
    notes: 'Zero Node.js fs dependencies. Ingest & export ArrayBuffer directly.',
    bootTime: '< 5 ms',
  },
  {
    name: 'Vercel Edge & Serverless',
    tier: 'Tier 1 - Edge / Serverless',
    status: 'Fully Supported',
    notes: 'Compatible with Next.js App Router route handlers. Instant cold starts.',
    bootTime: '< 20 ms',
  },
  {
    name: 'Bun (>= 1.0)',
    tier: 'Tier 1 - Alternate Runtime',
    status: 'Fully Supported',
    notes: 'High-speed execution leveraging Bun fast ZIP and Buffer APIs.',
    bootTime: '< 10 ms',
  },
  {
    name: 'Deno (>= 1.38)',
    tier: 'Tier 2 - Compatible',
    status: 'Supported via npm: specifier',
    notes: "Usable via import { PPTXTemplater } from 'npm:node-pptx-templater'.",
    bootTime: '< 25 ms',
  },
];

const CAPABILITIES_SLIDES = [
  {
    id: 'cap-1',
    category: 'Matrix',
    title: 'Runtime Compatibility Matrix',
    subtitle: 'Verified Across Serverless, Edge, and Modern JavaScript Engines',
    signature: 'Multi-Runtime Matrix',
    content: (
      <div className="space-y-3">
        <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)]">
          Pure JavaScript execution ensures 100% portability without architecture-specific native compilation:
        </p>
        <div className="overflow-x-auto ppt-card-surface rounded-lg border border-[var(--ppt-card-border)]">
          <table className="w-full text-left text-xs border-collapse">
            <thead>
              <tr className="border-b border-[var(--ppt-card-border)] text-[var(--ppt-text-secondary)]">
                <th className="py-2.5 px-3 font-semibold">Environment</th>
                <th className="py-2.5 px-3 font-semibold">Tier</th>
                <th className="py-2.5 px-3 font-semibold">Cold Start</th>
                <th className="py-2.5 px-3 font-semibold">Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-[var(--ppt-card-border)]/40">
              {RUNTIME_MATRIX.map((item, idx) => (
                <tr key={idx} className="hover:bg-black/5 dark:hover:bg-white/5">
                  <td className="py-2 px-3 font-medium text-[var(--ppt-text-primary)]">
                    {item.name}
                  </td>
                  <td className="py-2 px-3 text-[11px] text-[var(--ppt-text-secondary)]">
                    {item.tier}
                  </td>
                  <td className="py-2 px-3 font-mono text-[#D24726] font-bold">
                    {item.bootTime}
                  </td>
                  <td className="py-2 px-3">
                    <span className="inline-flex items-center gap-1 text-[11px] text-emerald-500 font-medium">
                      <CheckCircle2 className="w-3 h-3" />
                      {item.status}
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    ),
  },
  {
    id: 'cap-2',
    category: 'Architecture',
    title: 'Core Architectural Guarantees',
    subtitle: 'High-Throughput In-Memory Pipeline Designed for Enterprise Workloads',
    signature: 'Platform Guarantees',
    content: (
      <div className="space-y-3">
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div className="ppt-card-surface p-3.5 rounded-lg border border-[var(--ppt-card-border)]">
            <h4 className="text-xs font-bold text-[#D24726] mb-1">Zero Office / Electron Daemon</h4>
            <p className="text-[11px] text-[var(--ppt-text-secondary)] leading-relaxed">
              Unlike unoconv, LibreOffice, or Puppeteer, node-pptx-templater does not launch background daemons or require 500MB+ Docker containers.
            </p>
          </div>

          <div className="ppt-card-surface p-3.5 rounded-lg border border-[var(--ppt-card-border)]">
            <h4 className="text-xs font-bold text-[#D24726] mb-1">In-Memory Virtual File System</h4>
            <p className="text-[11px] text-[var(--ppt-text-secondary)] leading-relaxed">
              All OPC archive modifications are kept in an internal Map&lt;string, Uint8Array&gt;. Disk access is 100% optional.
            </p>
          </div>

          <div className="ppt-card-surface p-3.5 rounded-lg border border-[var(--ppt-card-border)]">
            <h4 className="text-xs font-bold text-[#D24726] mb-1">Linear Streaming I/O</h4>
            <p className="text-[11px] text-[var(--ppt-text-secondary)] leading-relaxed">
              Stream directly to HTTP clients, S3 buckets, or disk with minimal peak memory usage even for 500+ slide decks.
            </p>
          </div>

          <div className="ppt-card-surface p-3.5 rounded-lg border border-[var(--ppt-card-border)]">
            <h4 className="text-xs font-bold text-[#D24726] mb-1">Full OPC & Flat OPC Support</h4>
            <p className="text-[11px] text-[var(--ppt-text-secondary)] leading-relaxed">
              Native support for both standard PKZip-based .pptx presentations and single-file Flat OPC .xml format.
            </p>
          </div>
        </div>
      </div>
    ),
  },
  {
    id: 'cap-3',
    category: 'Frameworks',
    title: 'Direct Web Framework Adapters',
    subtitle: 'Stream PPTX Directly Through Express and Fastify',
    signature: 'Framework I/O',
    content: (
      <div className="space-y-3">
        <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)]">
          Send generated presentations with correct OpenXML MIME headers in one line of code:
        </p>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
          <div>
            <div className="text-xs font-bold text-[var(--ppt-text-primary)] mb-1">Express Route</div>
            <CodeBlock
              code={`app.get('/download-report', async (req, res) => {
  const ppt = await PPTXTemplater.load('./template.pptx');
  ppt.useSlide(1).replaceText({ '{{USER}}': req.user.name });
  await ppt.toExpress(res, 'Report.pptx');
});`}
              language="javascript"
              filename="express.js"
            />
          </div>

          <div>
            <div className="text-xs font-bold text-[var(--ppt-text-primary)] mb-1">Fastify Route</div>
            <CodeBlock
              code={`fastify.get('/download-report', async (req, reply) => {
  const ppt = await PPTXTemplater.load('./template.pptx');
  ppt.useSlide(1).replaceText({ '{{USER}}': req.user.name });
  await ppt.toFastify(reply, 'Report.pptx');
});`}
              language="javascript"
              filename="fastify.js"
            />
          </div>
        </div>

        <div className="flex justify-end pt-2">
          <Link
            href="/docs/api-reference"
            className="inline-flex items-center gap-1.5 px-4 py-2 rounded text-xs font-bold bg-[#D24726] text-white hover:bg-[#b83b1e] transition-colors"
          >
            <span>Next: Complete API Reference</span>
            <ArrowRight className="w-3.5 h-3.5" />
          </Link>
        </div>
      </div>
    ),
  },
];

export default function CapabilitiesPage() {
  return (
    <PowerPointSlideDeck
      slides={CAPABILITIES_SLIDES}
      categoryTitle="Capabilities"
    />
  );
}
