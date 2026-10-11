import React from 'react';
import Link from 'next/link';
import { ArrowRight, CheckCircle2, Cpu, ShieldCheck, Zap, Server, Globe } from 'lucide-react';
import compatibilityData from '../../../data/compatibility.json';
import CodeBlock from '../../../components/CodeBlock';

export const metadata = {
  title: 'Runtime Compatibility & Serverless Execution',
  description: 'Verified support across Node.js, AWS Lambda, Cloudflare Workers, and Vercel Edge.',
};

const LAMBDA_CODE = `import { PPTXTemplater } from 'node-pptx-templater';
import { S3Client, GetObjectCommand, PutObjectCommand } from '@aws-sdk/client-s3';

const s3 = new S3Client({ region: 'us-east-1' });

export const handler = async (event) => {
  // 1. Fetch template buffer from S3 bucket
  const s3Res = await s3.send(new GetObjectCommand({
    Bucket: 'company-templates',
    Key: 'executive-summary.pptx'
  }));
  const templateBuffer = Buffer.from(await s3Res.Body.transformToByteArray());

  // 2. Load into memory (zero /tmp disk I/O)
  const ppt = await PPTXTemplater.load(templateBuffer);
  ppt.useSlide(1).replaceText(event.replacements);

  // 3. Export directly as Node.js Buffer
  const outputBuffer = await ppt.toBuffer();

  // 4. Save generated deck to output bucket
  await s3.send(new PutObjectCommand({
    Bucket: 'company-reports',
    Key: \`reports/\${event.reportId}.pptx\`,
    Body: outputBuffer,
    ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
  }));

  return { statusCode: 200, body: JSON.stringify({ success: true }) };
};`;

const WORKERS_CODE = `import { PPTXTemplater } from 'node-pptx-templater';

export default {
  async fetch(request, env) {
    // 1. Fetch template from CDN or R2 bucket
    const res = await fetch('https://cdn.example.com/templates/invoice.pptx');
    const arrayBuffer = await res.arrayBuffer();

    // 2. Load in V8 isolate memory (no Node.js fs required)
    const ppt = await PPTXTemplater.load(arrayBuffer);
    ppt.useSlide(1).replaceText({ '{{CLIENT}}': 'Acme Global Corp' });

    // 3. Export ArrayBuffer directly to HTTP stream response
    const outputArrayBuffer = await ppt.toArrayBuffer();

    return new Response(outputArrayBuffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': 'attachment; filename="Invoice.pptx"'
      }
    });
  }
};`;

export default function CompatibilityPage() {
  return (
    <article className="space-y-12">
      {/* Page Header */}
      <div>
        <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full glass-pill text-xs font-semibold mb-4">
          <Cpu className="w-3.5 h-3.5" />
          <span>Serverless &amp; Edge Compatibility</span>
        </div>
        <h1 className="text-3xl sm:text-4xl font-extrabold tracking-tight text-slate-900 dark:text-white">
          Runtime Compatibility Matrix
        </h1>
        <p className="mt-3 text-base sm:text-lg text-slate-600 dark:text-slate-300 leading-relaxed">
          {compatibilityData.philosophy}
        </p>
      </div>

      {/* Runtime Compatibility Table */}
      <section className="space-y-4">
        <h2 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900 dark:text-white flex items-center gap-2">
          <Server className="w-5 h-5 text-brand-500" />
          <span>Supported Runtimes &amp; Environments</span>
        </h2>

        <div className="overflow-x-auto rounded-2xl glass-card">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-100/80 dark:bg-slate-950/80 uppercase font-mono text-[10px] text-slate-500 dark:text-slate-400 border-b border-slate-200 dark:border-slate-800">
              <tr>
                <th className="px-4 py-3">Environment / Runtime</th>
                <th className="px-4 py-3">Status</th>
                <th className="px-4 py-3">Cold Boot</th>
                <th className="px-4 py-3">Package Size</th>
                <th className="px-4 py-3">Details</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 dark:divide-slate-800/60 font-sans">
              {compatibilityData.runtimes.map((runtime, idx) => (
                <tr key={idx} className="hover:bg-slate-50/50 dark:hover:bg-slate-900/40">
                  <td className="px-4 py-3.5 font-semibold text-slate-900 dark:text-white whitespace-nowrap">
                    {runtime.name}
                  </td>
                  <td className="px-4 py-3.5 whitespace-nowrap">
                    <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border border-emerald-500/20">
                      <CheckCircle2 className="w-3 h-3" />
                      {runtime.status}
                    </span>
                  </td>
                  <td className="px-4 py-3.5 font-mono text-brand-600 dark:text-brand-400 whitespace-nowrap">
                    {runtime.bootTime}
                  </td>
                  <td className="px-4 py-3.5 font-mono text-slate-600 dark:text-slate-400 whitespace-nowrap">
                    {runtime.bundleSize}
                  </td>
                  <td className="px-4 py-3.5 text-slate-600 dark:text-slate-400 min-w-[220px]">
                    {runtime.notes}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      {/* Production Guarantees */}
      <section className="space-y-4">
        <h2 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900 dark:text-white flex items-center gap-2">
          <ShieldCheck className="w-5 h-5 text-emerald-500" />
          <span>Production Guarantees</span>
        </h2>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          {compatibilityData.guarantees.map((g, idx) => (
            <div key={idx} className="glass-card p-5 rounded-2xl space-y-2">
              <h4 className="text-sm font-bold text-slate-900 dark:text-white flex items-center gap-2">
                <CheckCircle2 className="w-4 h-4 text-brand-500" />
                <span>{g.title}</span>
              </h4>
              <p className="text-xs text-slate-600 dark:text-slate-400 leading-relaxed">
                {g.description}
              </p>
            </div>
          ))}
        </div>
      </section>

      {/* Code Examples */}
      <section className="space-y-6">
        <h2 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900 dark:text-white">
          Deployment Recipes
        </h2>

        <div className="space-y-3">
          <h3 className="text-sm font-bold text-slate-900 dark:text-white flex items-center gap-2">
            <span>AWS Lambda: S3 Streaming &amp; In-Memory Buffer</span>
          </h3>
          <CodeBlock code={LAMBDA_CODE} language="javascript" filename="lambda-handler.mjs" />
        </div>

        <div className="space-y-3">
          <h3 className="text-sm font-bold text-slate-900 dark:text-white flex items-center gap-2">
            <span>Cloudflare Workers: V8 Isolate Streaming</span>
          </h3>
          <CodeBlock code={WORKERS_CODE} language="javascript" filename="worker.js" />
        </div>
      </section>

      {/* Next Step CTA */}
      <div className="pt-8 border-t border-slate-200 dark:border-slate-800 flex flex-wrap items-center justify-between gap-4">
        <div>
          <span className="text-xs text-slate-500">Next Step</span>
          <p className="text-sm font-semibold text-slate-900 dark:text-white">Explore the Complete Methods API Catalog</p>
        </div>
        <Link
          href="/docs/api"
          className="flex items-center gap-2 px-5 py-2.5 rounded-xl text-xs font-semibold bg-brand-600 text-white hover:bg-brand-500 shadow-md shadow-brand-500/20 transition-all"
        >
          <span>API Reference</span>
          <ArrowRight className="w-3.5 h-3.5" />
        </Link>
      </div>
    </article>
  );
}
