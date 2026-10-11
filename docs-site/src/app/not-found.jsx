import React from 'react';
import Link from 'next/link';
import { FileQuestion, ArrowRight, Home, Code2 } from 'lucide-react';

export default function NotFound() {
  return (
    <div className="max-w-xl mx-auto px-4 py-24 text-center">
      <div className="w-16 h-16 rounded-2xl bg-brand-500/10 border border-brand-500/20 flex items-center justify-center mx-auto mb-6 text-brand-400">
        <FileQuestion className="w-8 h-8" />
      </div>

      <h1 className="text-4xl font-extrabold text-white tracking-tight">
        404 — Page Not Found
      </h1>

      <p className="mt-4 text-sm text-slate-400 leading-relaxed">
        The documentation page or API endpoint you requested does not exist or has been relocated in the v3.0 modernization.
      </p>

      <div className="mt-8 flex flex-wrap items-center justify-center gap-3">
        <Link
          href="/"
          className="px-5 py-2.5 rounded-xl font-semibold text-xs bg-brand-500 hover:bg-brand-400 text-white flex items-center gap-2 transition-colors"
        >
          <Home className="w-4 h-4" />
          <span>Back to Home</span>
        </Link>
        <Link
          href="/api-reference/"
          className="px-5 py-2.5 rounded-xl font-semibold text-xs bg-slate-900 hover:bg-slate-800 text-slate-200 border border-slate-700/80 flex items-center gap-2 transition-colors"
        >
          <Code2 className="w-4 h-4 text-brand-400" />
          <span>Browse 101 APIs</span>
        </Link>
      </div>
    </div>
  );
}
