import React from 'react';
import Link from 'next/link';
import { FileQuestion, ArrowRight, Home, Code2 } from 'lucide-react';

export default function NotFound() {
  return (
    <div className="max-w-xl mx-auto px-4 py-24 text-center">
      <div className="w-16 h-16 rounded-2xl bg-[#D24726]/10 border border-[#D24726]/30 flex items-center justify-center mx-auto mb-6 text-[#D24726]">
        <FileQuestion className="w-8 h-8" />
      </div>

      <h1 className="text-3xl font-extrabold text-[var(--ppt-text-primary)] tracking-tight">
        404 — Slide Not Found
      </h1>

      <p className="mt-4 text-xs sm:text-sm text-[var(--ppt-text-secondary)] leading-relaxed">
        The requested presentation topic or route does not exist. Please use the PowerPoint ribbon above to navigate to our documentation sections.
      </p>

      <div className="mt-8 flex flex-wrap items-center justify-center gap-3">
        <Link
          href="/"
          className="px-5 py-2.5 rounded-lg font-semibold text-xs bg-[#D24726] hover:bg-[#b83b1e] text-white flex items-center gap-2 transition-colors shadow-md"
        >
          <Home className="w-4 h-4" />
          <span>Home Presentation</span>
        </Link>
        <Link
          href="/docs/api-reference"
          className="px-5 py-2.5 rounded-lg font-semibold text-xs ppt-card-surface text-[var(--ppt-text-primary)] hover:border-[#D24726] flex items-center gap-2 transition-colors"
        >
          <Code2 className="w-4 h-4 text-[#D24726]" />
          <span>API Reference Catalog</span>
        </Link>
      </div>
    </div>
  );
}
