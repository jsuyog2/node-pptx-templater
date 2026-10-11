import React, { useState } from 'react';
import { Hash, Check, BookOpen, Layers, Terminal } from 'lucide-react';
import CodeBlock from './CodeBlock';

export default function ApiCard({ method }) {
  const [copiedLink, setCopiedLink] = useState(false);
  const anchorId = `api-${method.name}`;

  const copyAnchor = (e) => {
    e.preventDefault();
    const url = `${window.location.origin}${window.location.pathname}#${anchorId}`;
    navigator.clipboard.writeText(url);
    window.location.hash = anchorId;
    setCopiedLink(true);
    setTimeout(() => setCopiedLink(false), 2000);
  };

  return (
    <article
      id={anchorId}
      className="glass-card rounded-2xl p-6 sm:p-7 mb-8 scroll-mt-24 transition-all duration-200"
    >
      {/* Header Bar */}
      <div className="flex flex-wrap items-center justify-between gap-3 pb-4 border-b border-slate-200/80 dark:border-slate-800/80">
        <div className="flex items-center gap-2.5">
          <a
            href={`#${anchorId}`}
            onClick={copyAnchor}
            title="Copy direct link to this method"
            className="p-1 text-slate-400 hover:text-blue-600 dark:hover:text-blue-400 transition-colors rounded"
          >
            {copiedLink ? (
              <Check className="w-4 h-4 text-emerald-500" />
            ) : (
              <Hash className="w-4 h-4" />
            )}
          </a>
          <h3 className="text-lg sm:text-xl font-bold font-mono text-slate-900 dark:text-slate-100 tracking-tight">
            {method.name}
          </h3>
        </div>

        {/* Badges */}
        <div className="flex flex-wrap items-center gap-1.5">
          {method.isStatic && (
            <span className="px-2.5 py-0.5 rounded-full text-xs font-semibold bg-amber-500/10 text-amber-600 dark:text-amber-400 border border-amber-500/20">
              static
            </span>
          )}
          {method.isAsync && (
            <span className="px-2.5 py-0.5 rounded-full text-xs font-semibold bg-purple-500/10 text-purple-600 dark:text-purple-400 border border-purple-500/20">
              async
            </span>
          )}
          <span className="px-2.5 py-0.5 rounded-full text-xs font-medium bg-blue-500/10 text-blue-600 dark:text-blue-400 border border-blue-500/20">
            {method.category}
          </span>
          {method.returns?.type && (
            <span className="px-2 py-0.5 rounded-full text-xs font-mono bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300 border border-slate-200 dark:border-slate-700">
              → {method.returns.type}
            </span>
          )}
        </div>
      </div>

      {/* Signature */}
      <div className="my-4 px-3.5 py-2.5 rounded-lg bg-slate-100/80 dark:bg-slate-950/60 border border-slate-200/80 dark:border-slate-800/80 font-mono text-xs sm:text-sm text-blue-700 dark:text-blue-300 overflow-x-auto">
        <code>{method.signature}</code>
      </div>

      {/* Description */}
      <p className="text-sm sm:text-base text-slate-600 dark:text-slate-300 leading-relaxed my-3">
        {method.description}
      </p>

      {/* OpenXML Note Callout */}
      {/* {method.openXml && (
        <div className="flex items-start gap-2.5 my-3 p-3 rounded-xl bg-slate-50 dark:bg-slate-900/40 border border-slate-200 dark:border-slate-800/60 text-xs sm:text-sm text-slate-600 dark:text-slate-400">
          <Layers className="w-4 h-4 text-blue-500 flex-shrink-0 mt-0.5" />
          <div>
            <strong className="text-slate-700 dark:text-slate-300 font-semibold">OpenXML Specification: </strong>
            <span>{method.openXml}</span>
          </div>
        </div>
      )} */}

      {/* Parameters Table */}
      {method.parameters && method.parameters.length > 0 && (
        <div className="mt-5 mb-4">
          <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400 dark:text-slate-500 mb-2 flex items-center gap-1.5">
            <BookOpen className="w-3.5 h-3.5" /> Parameters
          </h4>
          <div className="overflow-x-auto border border-slate-200 dark:border-slate-800 rounded-xl">
            <table className="w-full text-left text-xs sm:text-sm border-collapse">
              <thead>
                <tr className="bg-slate-100/70 dark:bg-slate-800/50 border-b border-slate-200 dark:border-slate-800 text-slate-600 dark:text-slate-300 font-semibold">
                  <th className="py-2.5 px-3">Name</th>
                  <th className="py-2.5 px-3">Type</th>
                  <th className="py-2.5 px-3">Required</th>
                  <th className="py-2.5 px-3">Default</th>
                  <th className="py-2.5 px-3">Description</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200/60 dark:divide-slate-800/60 text-slate-700 dark:text-slate-300">
                {method.parameters.map((param, idx) => (
                  <tr key={idx} className="hover:bg-slate-50/50 dark:hover:bg-slate-800/30">
                    <td className="py-2.5 px-3 font-mono font-medium text-blue-600 dark:text-blue-400 whitespace-nowrap">
                      {param.name}
                    </td>
                    <td className="py-2.5 px-3 font-mono text-xs text-purple-600 dark:text-purple-300 whitespace-nowrap">
                      {param.type}
                    </td>
                    <td className="py-2.5 px-3 whitespace-nowrap">
                      {param.required ? (
                        <span className="text-xs font-semibold text-rose-500 dark:text-rose-400">Yes</span>
                      ) : (
                        <span className="text-xs text-slate-400">No</span>
                      )}
                    </td>
                    <td className="py-2.5 px-3 font-mono text-xs text-slate-500 whitespace-nowrap">
                      {param.default}
                    </td>
                    <td className="py-2.5 px-3 leading-relaxed text-xs sm:text-sm">
                      {param.description}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Returns Section */}
      {method.returns && (
        <div className="mt-3 mb-4 text-xs sm:text-sm">
          <span className="font-semibold text-slate-700 dark:text-slate-300">Returns: </span>
          <code className="text-purple-600 dark:text-purple-400 font-mono font-medium">
            {method.returns.type}
          </code>
          {method.returns.description && (
            <span className="text-slate-600 dark:text-slate-400 ml-2">
              — {method.returns.description}
            </span>
          )}
        </div>
      )}

      {/* Code Snippet */}
      {method.exampleCode && (
        <div className="mt-4">
          <CodeBlock
            code={method.exampleCode}
            language="javascript"
            title={`Example: ${method.name}`}
          />
        </div>
      )}
    </article>
  );
}
