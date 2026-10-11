'use client';

import React, { useState } from 'react';
import { Check, Copy } from 'lucide-react';

export default function CodeBlock({ code = '', language = 'javascript', filename = '' }) {
  const [copied, setCopied] = useState(false);

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(code);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      console.error('Failed to copy code: ', err);
    }
  };

  return (
    <div className="relative group my-4 rounded-xl border border-slate-700/60 bg-slate-900/95 overflow-hidden shadow-xl shadow-black/30">
      <div className="flex items-center justify-between px-4 py-2 border-b border-slate-800 bg-slate-950/70 text-xs font-mono text-slate-400">
        <div className="flex items-center gap-2">
          <div className="flex gap-1.5">
            <span className="w-2.5 h-2.5 rounded-full bg-rose-500/80 inline-block"></span>
            <span className="w-2.5 h-2.5 rounded-full bg-amber-500/80 inline-block"></span>
            <span className="w-2.5 h-2.5 rounded-full bg-emerald-500/80 inline-block"></span>
          </div>
          {filename && <span className="ml-2 font-medium text-slate-300">{filename}</span>}
          {!filename && <span className="uppercase text-slate-500 tracking-wider">{language}</span>}
        </div>
        <button
          onClick={handleCopy}
          aria-label="Copy code snippet"
          className="flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-medium bg-slate-800/90 hover:bg-slate-700 text-slate-300 hover:text-white transition-all duration-150 border border-slate-700/50"
        >
          {copied ? (
            <>
              <Check className="w-3.5 h-3.5 text-emerald-400" />
              <span className="text-emerald-400 font-sans">Copied!</span>
            </>
          ) : (
            <>
              <Copy className="w-3.5 h-3.5 text-slate-400" />
              <span className="font-sans">Copy</span>
            </>
          )}
        </button>
      </div>
      <div className="p-4 overflow-x-auto text-sm text-slate-200 leading-relaxed font-mono selection:bg-brand-500/30 selection:text-brand-200">
        <pre className="whitespace-pre">
          <code>{code}</code>
        </pre>
      </div>
    </div>
  );
}
