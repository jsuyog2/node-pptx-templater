import React from 'react';
import Link from 'next/link';
import { Github, Heart, Layers } from 'lucide-react';
import navigation from '../data/navigation.json';
import siteMeta from '../data/siteMeta.json';

export default function Footer() {
  return (
    <footer className="border-t border-slate-800/80 bg-slate-950 text-slate-400 text-sm mt-20">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-12">
        <div className="grid grid-cols-1 md:grid-cols-4 gap-8">
          <div className="md:col-span-1 space-y-3">
            <div className="flex items-center gap-2.5">
              <div className="w-7 h-7 rounded-lg bg-brand-500/20 border border-brand-500/40 flex items-center justify-center">
                <Layers className="w-4 h-4 text-brand-400" />
              </div>
              <span className="font-bold text-white tracking-tight">node-pptx-templater</span>
            </div>
            <p className="text-xs text-slate-400 leading-relaxed">
              {siteMeta.tagline}
            </p>
            <div className="text-xs text-slate-500">
              Released under {siteMeta.license} License.
            </div>
          </div>

          <div>
            <h4 className="text-xs font-semibold text-slate-200 uppercase tracking-wider mb-3">Documentation</h4>
            <ul className="space-y-2 text-xs">
              <li>
                <Link href="/" className="hover:text-white transition-colors">Overview</Link>
              </li>
              <li>
                <Link href="/api-reference/" className="hover:text-white transition-colors">API Reference (10 Domains)</Link>
              </li>
              <li>
                <Link href="/guides/quick-start/" className="hover:text-white transition-colors">Quick Start Guide</Link>
              </li>
              <li>
                <Link href="/guides/serverless-deployment/" className="hover:text-white transition-colors">Serverless Runtimes</Link>
              </li>
            </ul>
          </div>

          <div>
            <h4 className="text-xs font-semibold text-slate-200 uppercase tracking-wider mb-3">Deep Dives</h4>
            <ul className="space-y-2 text-xs">
              <li>
                <Link href="/guides/run-fragment-healing/" className="hover:text-white transition-colors">Run Fragment Healing</Link>
              </li>
              <li>
                <Link href="/guides/chart-dual-sync/" className="hover:text-white transition-colors">Chart Excel Dual-Sync</Link>
              </li>
              <li>
                <Link href="/guides/rectangular-table-spans/" className="hover:text-white transition-colors">Rectangular Table Spans</Link>
              </li>
            </ul>
          </div>

          <div>
            <h4 className="text-xs font-semibold text-slate-200 uppercase tracking-wider mb-3">Community & Ecosystem</h4>
            <ul className="space-y-2 text-xs">
              <li>
                <a href={navigation.external.github} target="_blank" rel="noopener noreferrer" className="hover:text-white transition-colors flex items-center gap-1.5">
                  <Github className="w-3.5 h-3.5" />
                  GitHub Repository
                </a>
              </li>
              <li>
                <a href={navigation.external.npm} target="_blank" rel="noopener noreferrer" className="hover:text-white transition-colors">
                  npm Package ({siteMeta.version})
                </a>
              </li>
              <li>
                <a href={navigation.external.releases} target="_blank" rel="noopener noreferrer" className="hover:text-white transition-colors">
                  Changelog & Releases
                </a>
              </li>
              <li>
                <a href={navigation.external.issues} target="_blank" rel="noopener noreferrer" className="hover:text-white transition-colors">
                  Issue Tracker
                </a>
              </li>
            </ul>
          </div>
        </div>

        <div className="mt-8 pt-8 border-t border-slate-900 flex flex-col sm:flex-row items-center justify-between text-xs text-slate-500 gap-4">
          <p>© {new Date().getFullYear()} {siteMeta.author}. Pure JavaScript OpenXML Engine.</p>
          <div className="flex items-center gap-1 text-slate-400">
            Built with Next.js static export &amp; hosted on Cloudflare Pages
          </div>
        </div>
      </div>
    </footer>
  );
}
