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
    <div className="relative group my-4 text-start rounded-lg border border-[var(--ppt-card-border)] bg-black/40 dark:bg-black/60 overflow-hidden shadow-md">
      {/* Code Header Bar */}
      <div className="flex items-center justify-between px-3.5 py-1.5 border-b border-[var(--ppt-card-border)] bg-black/20 dark:bg-black/40 text-xs font-mono text-[var(--ppt-text-secondary)]">
        <div className="flex items-center gap-2">
          <div className="flex gap-1.5">
            <span className="w-2.5 h-2.5 rounded-full bg-[#D24726] inline-block opacity-80" />
            <span className="w-2.5 h-2.5 rounded-full bg-amber-500/80 inline-block" />
            <span className="w-2.5 h-2.5 rounded-full bg-emerald-500/80 inline-block" />
          </div>
          {filename && <span className="ml-2 font-medium text-[var(--ppt-text-primary)]">{filename}</span>}
          {!filename && <span className="uppercase text-[10px] tracking-wider text-[var(--ppt-text-secondary)]">{language}</span>}
        </div>

        <button
          onClick={handleCopy}
          type="button"
          aria-label="Copy code snippet"
          className="flex items-center gap-1.5 px-2 py-0.5 rounded text-[11px] font-sans font-medium bg-black/30 hover:bg-[#D24726] text-[var(--ppt-text-secondary)] hover:text-white transition-all border border-[var(--ppt-card-border)] hover:border-[#D24726]"
        >
          {copied ? (
            <>
              <Check className="w-3 h-3 text-emerald-400" />
              <span className="text-emerald-400 font-sans">Copied!</span>
            </>
          ) : (
            <>
              <Copy className="w-3 h-3" />
              <span>Copy</span>
            </>
          )}
        </button>
      </div>

      {/* Code Contents */}
      <div
        onTouchStart={(e) => e.stopPropagation()}
        onTouchMove={(e) => e.stopPropagation()}
        onTouchEnd={(e) => e.stopPropagation()}
        className="p-4 overflow-x-auto text-xs sm:text-sm text-[var(--ppt-text-primary)] leading-relaxed font-mono selection:bg-[#D24726]/30 selection:text-white"
      >
        <pre className="whitespace-pre">
          <code>{code}</code>
        </pre>
      </div>
    </div>
  );
}
