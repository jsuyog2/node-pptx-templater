import React from 'react';
import { Github, Box, Heart, ExternalLink, ShieldCheck, Zap } from 'lucide-react';

export default function Footer() {
  return (
    <footer className="mt-20 border-t border-slate-200/80 dark:border-slate-800/80 bg-white/50 dark:bg-slate-950/40 backdrop-blur-md">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-12">
        <div className="grid grid-cols-1 md:grid-cols-4 gap-8 mb-8">
          {/* Col 1: Brand & Mandate */}
          <div className="md:col-span-2">
            <div className="flex items-center gap-2.5 mb-3">
              <div className="w-7 h-7 rounded-lg bg-blue-600 flex items-center justify-center text-white font-black text-xs shadow-sm">
                P
              </div>
              <span className="font-bold text-base tracking-tight text-slate-900 dark:text-slate-100">
                node-pptx-templater
              </span>
              <span className="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border border-emerald-500/20">
                v3.0.0-alpha.0
              </span>
            </div>
            <p className="text-xs sm:text-sm text-slate-500 dark:text-slate-400 max-w-md leading-relaxed mb-4">
              Pure JavaScript PowerPoint (PPTX) OpenXML manipulation and template engine for Node.js, AWS Lambda, Cloudflare Workers, and Vercel Edge. Zero Microsoft Office or Java dependencies.
            </p>
            <div className="flex items-center gap-4 text-xs text-slate-500 dark:text-slate-400">
              <span className="flex items-center gap-1 text-emerald-600 dark:text-emerald-400">
                <ShieldCheck className="w-3.5 h-3.5" /> 285+ Vitest Tests Passing
              </span>
              <span className="flex items-center gap-1 text-blue-600 dark:text-blue-400">
                <Zap className="w-3.5 h-3.5" /> Pure In-Memory VFS
              </span>
            </div>
          </div>

          {/* Col 2: Documentation */}
          <div>
            <h4 className="text-xs font-bold uppercase tracking-wider text-slate-900 dark:text-slate-100 mb-3">
              Documentation
            </h4>
            <ul className="space-y-2 text-xs text-slate-500 dark:text-slate-400">
              <li>
                <a href="#/guides/getting-started" className="hover:text-blue-600 dark:hover:text-blue-400 transition-colors">
                  Installation & Setup
                </a>
              </li>
              <li>
                <a href="#/guides/serverless" className="hover:text-blue-600 dark:hover:text-blue-400 transition-colors">
                  Serverless (AWS, Cloudflare, Edge)
                </a>
              </li>
              <li>
                <a href="#/guides/excel-dual-sync" className="hover:text-blue-600 dark:hover:text-blue-400 transition-colors">
                  Excel Dual-Sync Architecture
                </a>
              </li>
              <li>
                <a href="#/guides/table-merging" className="hover:text-blue-600 dark:hover:text-blue-400 transition-colors">
                  Table Pagination & Merging
                </a>
              </li>
              <li>
                <a href="#/api-reference" className="hover:text-blue-600 dark:hover:text-blue-400 transition-colors">
                  Full API Reference (102 Methods)
                </a>
              </li>
            </ul>
          </div>

          {/* Col 3: Ecosystem & Community */}
          <div>
            <h4 className="text-xs font-bold uppercase tracking-wider text-slate-900 dark:text-slate-100 mb-3">
              Community & Code
            </h4>
            <ul className="space-y-2 text-xs text-slate-500 dark:text-slate-400">
              <li>
                <a
                  href="https://github.com/jsuyog2/node-pptx-templater"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="flex items-center gap-1.5 hover:text-blue-600 dark:hover:text-blue-400 transition-colors"
                >
                  <Github className="w-3.5 h-3.5" /> GitHub Repository
                </a>
              </li>
              <li>
                <a
                  href="https://www.npmjs.com/package/node-pptx-templater"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="flex items-center gap-1.5 hover:text-blue-600 dark:hover:text-blue-400 transition-colors"
                >
                  <Box className="w-3.5 h-3.5 text-rose-500" /> npm Package
                </a>
              </li>
              <li>
                <a
                  href="https://github.com/jsuyog2/node-pptx-templater/issues"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="hover:text-blue-600 dark:hover:text-blue-400 transition-colors"
                >
                  Issue Tracker & Bugs
                </a>
              </li>
              <li>
                <a
                  href="https://opensource.org/licenses/MIT"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="hover:text-blue-600 dark:hover:text-blue-400 transition-colors"
                >
                  MIT License
                </a>
              </li>
            </ul>
          </div>
        </div>

        {/* Bottom Bar */}
        <div className="pt-8 border-t border-slate-200/60 dark:border-slate-800/60 flex flex-col sm:flex-row items-center justify-between gap-4 text-xs text-slate-500 dark:text-slate-400">
          <div>
            © {new Date().getFullYear()} node-pptx-templater contributors. Released under the MIT License.
          </div>
          <div className="flex items-center gap-1">
            Built for developers automating enterprise slides at scale.
          </div>
        </div>
      </div>
    </footer>
  );
}
