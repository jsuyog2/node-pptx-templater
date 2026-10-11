'use client';

import React from 'react';
import { Hash, AlertCircle, Cpu, Zap, ArrowRight, CheckCircle2 } from 'lucide-react';
import CodeBlock from './CodeBlock';

export default function ApiCard({ method }) {
  if (!method) return null;

  return (
    <article
      id={method.name}
      className="scroll-mt-24 p-6 rounded-2xl glass-card transition-all duration-200 hover:-translate-y-0.5"
    >
      {/* Header with Title, Badges, and Anchor */}
      <div className="flex flex-wrap items-center justify-between gap-3 pb-4 border-b border-slate-200 dark:border-slate-800/80">
        <div className="flex items-center gap-2.5">
          <a
            href={`#${method.name}`}
            className="text-slate-400 hover:text-brand-500 group p-1 -ml-1 transition-colors"
            title={`Anchor link for ${method.name}`}
            aria-label={`Link to ${method.name}`}
          >
            <Hash className="w-5 h-5 text-slate-400 group-hover:text-brand-500" />
          </a>
          <h3 className="text-xl font-bold font-mono text-slate-900 dark:text-white tracking-tight">
            {method.name}
          </h3>
          <span className="text-xs px-2.5 py-0.5 rounded-full font-semibold bg-brand-500/10 text-brand-600 dark:text-brand-400 border border-brand-500/20">
            {method.category}
          </span>
        </div>

        <div className="flex items-center gap-1.5 flex-wrap">
          {method.isStatic && (
            <span className="text-[11px] font-mono font-medium px-2 py-0.5 rounded bg-purple-500/15 text-purple-700 dark:text-purple-300 border border-purple-500/30">
              static
            </span>
          )}
          {method.isAsync && (
            <span className="text-[11px] font-mono font-medium px-2 py-0.5 rounded bg-amber-500/15 text-amber-700 dark:text-amber-300 border border-amber-500/30">
              async
            </span>
          )}
          <span className="text-[11px] font-mono font-medium px-2 py-0.5 rounded bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 border border-slate-200 dark:border-slate-700">
            returns {typeof method.returns === 'object' ? method.returns.type : method.returns}
          </span>
        </div>
      </div>

      {/* Signature Code Box */}
      <div className="my-3 px-3.5 py-2.5 rounded-xl bg-slate-900 dark:bg-slate-950 text-sm font-mono text-brand-300 overflow-x-auto shadow-inner border border-slate-800">
        <code>{method.signature}</code>
      </div>

      {/* Description */}
      <p className="text-sm text-slate-600 dark:text-slate-300 leading-relaxed my-3">
        {method.description}
      </p>

      {/* Parameters Table */}
      {method.parameters && method.parameters.length > 0 && (
        <div className="my-4">
          <h4 className="text-xs font-semibold uppercase tracking-wider text-slate-500 dark:text-slate-400 mb-2">
            Parameters
          </h4>
          <div className="overflow-x-auto rounded-xl border border-slate-200 dark:border-slate-800 bg-white/50 dark:bg-slate-950/50">
            <table className="w-full text-left text-xs text-slate-700 dark:text-slate-300">
              <thead className="bg-slate-100/80 dark:bg-slate-900/90 text-slate-500 dark:text-slate-400 uppercase font-mono text-[10px] border-b border-slate-200 dark:border-slate-800">
                <tr>
                  <th className="px-3.5 py-2">Name</th>
                  <th className="px-3.5 py-2">Type</th>
                  <th className="px-3.5 py-2">Required</th>
                  <th className="px-3.5 py-2">Default</th>
                  <th className="px-3.5 py-2">Description</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 dark:divide-slate-800/60 font-sans">
                {method.parameters.map((param, idx) => (
                  <tr key={idx} className="hover:bg-slate-50/50 dark:hover:bg-slate-900/40">
                    <td className="px-3.5 py-2 font-mono font-semibold text-brand-600 dark:text-brand-400 whitespace-nowrap">
                      {param.name}
                    </td>
                    <td className="px-3.5 py-2 font-mono text-purple-600 dark:text-purple-300 whitespace-nowrap">
                      {param.type}
                    </td>
                    <td className="px-3.5 py-2 whitespace-nowrap">
                      {param.required ? (
                        <span className="text-rose-500 font-semibold">Yes</span>
                      ) : (
                        <span className="text-slate-400">No</span>
                      )}
                    </td>
                    <td className="px-3.5 py-2 font-mono text-slate-500 dark:text-slate-400 whitespace-nowrap">
                      {param.default !== undefined ? String(param.default) : '—'}
                    </td>
                    <td className="px-3.5 py-2 text-slate-600 dark:text-slate-300 min-w-[200px]">
                      {param.description}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* OpenXML Impact & Edge Cases Callouts */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-3 my-4 text-xs">
        {method.openXmlImpact && (
          <div className="p-3.5 rounded-xl bg-indigo-50/70 dark:bg-indigo-950/20 border border-indigo-200 dark:border-indigo-900/40 text-indigo-900 dark:text-indigo-200">
            <div className="flex items-center gap-1.5 font-semibold text-indigo-700 dark:text-indigo-300 mb-1">
              <Cpu className="w-3.5 h-3.5" />
              <span>OpenXML Package Impact</span>
            </div>
            <p className="text-slate-600 dark:text-slate-300 leading-relaxed font-mono text-[11px]">
              {method.openXmlImpact}
            </p>
          </div>
        )}

        {method.edgeCases && (
          <div className="p-3.5 rounded-xl bg-amber-50/70 dark:bg-amber-950/20 border border-amber-200 dark:border-amber-900/40 text-amber-900 dark:text-amber-200">
            <div className="flex items-center gap-1.5 font-semibold text-amber-700 dark:text-amber-300 mb-1">
              <AlertCircle className="w-3.5 h-3.5" />
              <span>Edge Cases &amp; Diagnostics</span>
            </div>
            <p className="text-slate-600 dark:text-slate-300 leading-relaxed text-[11px]">
              {method.edgeCases}
            </p>
          </div>
        )}
      </div>

      {/* Copyable Example Code */}
      {method.exampleCode && (
        <div className="mt-4">
          <h4 className="text-xs font-semibold uppercase tracking-wider text-slate-500 dark:text-slate-400 mb-1">
            Usage Example
          </h4>
          <CodeBlock code={method.exampleCode} language="javascript" />
        </div>
      )}
    </article>
  );
}
