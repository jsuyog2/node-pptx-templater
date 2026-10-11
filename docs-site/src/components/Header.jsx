'use client';

import React, { useState } from 'react';
import Link from 'next/link';
import { BookOpen, Code2, ExternalLink, Github, Menu, X, Layers, Cpu } from 'lucide-react';
import navigation from '../data/navigation.json';
import siteMeta from '../data/siteMeta.json';

export default function Header() {
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);

  return (
    <header className="sticky top-0 z-50 backdrop-blur-md bg-slate-950/80 border-b border-slate-800/80">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex items-center justify-between h-16">
          {/* Logo & Brand */}
          <div className="flex items-center gap-3">
            <Link href="/" className="flex items-center gap-3 group">
              <div className="w-9 h-9 rounded-xl bg-gradient-to-tr from-brand-600 to-indigo-500 p-0.5 shadow-lg shadow-brand-500/20 group-hover:scale-105 transition-transform duration-200">
                <div className="w-full h-full bg-slate-950 rounded-[10px] flex items-center justify-center">
                  <Layers className="w-5 h-5 text-brand-400 group-hover:text-brand-300 transition-colors" />
                </div>
              </div>
              <div className="flex flex-col">
                <div className="flex items-center gap-2">
                  <span className="font-bold text-base sm:text-lg tracking-tight text-white group-hover:text-brand-300 transition-colors">
                    node-pptx-templater
                  </span>
                  <span className="hidden sm:inline-block text-[11px] font-semibold px-2 py-0.5 rounded-full bg-brand-500/15 text-brand-400 border border-brand-500/30">
                    {siteMeta.version}
                  </span>
                </div>
                <span className="text-[10px] text-slate-400 font-mono -mt-0.5 hidden sm:block">
                  Pure JS OpenXML Template Engine
                </span>
              </div>
            </Link>
          </div>

          {/* Desktop Navigation */}
          <nav className="hidden md:flex items-center gap-1">
            <Link
              href="/"
              className="px-3.5 py-1.5 rounded-lg text-sm font-medium text-slate-300 hover:text-white hover:bg-slate-800/60 transition-colors"
            >
              Overview
            </Link>
            <Link
              href="/api-reference/"
              className="px-3.5 py-1.5 rounded-lg text-sm font-medium text-slate-300 hover:text-white hover:bg-slate-800/60 transition-colors flex items-center gap-1.5"
            >
              <Code2 className="w-4 h-4 text-brand-400" />
              API Reference
            </Link>
            <Link
              href="/guides/quick-start/"
              className="px-3.5 py-1.5 rounded-lg text-sm font-medium text-slate-300 hover:text-white hover:bg-slate-800/60 transition-colors flex items-center gap-1.5"
            >
              <BookOpen className="w-4 h-4 text-emerald-400" />
              Guides
            </Link>
            <Link
              href="/guides/serverless-deployment/"
              className="px-3.5 py-1.5 rounded-lg text-sm font-medium text-slate-300 hover:text-white hover:bg-slate-800/60 transition-colors flex items-center gap-1.5"
            >
              <Cpu className="w-4 h-4 text-indigo-400" />
              Serverless
            </Link>
          </nav>

          {/* Right Action Buttons */}
          <div className="hidden sm:flex items-center gap-3">
            <a
              href={navigation.external.github}
              target="_blank"
              rel="noopener noreferrer"
              className="flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-semibold bg-slate-900 hover:bg-slate-800 text-slate-200 border border-slate-700/60 transition-all shadow-sm"
            >
              <Github className="w-4 h-4" />
              <span>GitHub</span>
            </a>
            <a
              href={navigation.external.npm}
              target="_blank"
              rel="noopener noreferrer"
              className="flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-semibold bg-rose-950/40 hover:bg-rose-900/50 text-rose-300 border border-rose-800/40 transition-all shadow-sm"
            >
              <span className="font-mono font-bold">npm</span>
              <ExternalLink className="w-3.5 h-3.5" />
            </a>
          </div>

          {/* Mobile Menu Button */}
          <div className="flex md:hidden items-center gap-2">
            <button
              onClick={() => setMobileMenuOpen(!mobileMenuOpen)}
              className="p-2 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 focus:outline-none"
              aria-label="Toggle navigation menu"
            >
              {mobileMenuOpen ? <X className="w-6 h-6" /> : <Menu className="w-6 h-6" />}
            </button>
          </div>
        </div>
      </div>

      {/* Mobile Drawer */}
      {mobileMenuOpen && (
        <div className="md:hidden border-b border-slate-800 bg-slate-950/95 px-4 pt-2 pb-6 space-y-2 backdrop-blur-xl">
          <Link
            href="/"
            onClick={() => setMobileMenuOpen(false)}
            className="block px-3 py-2 rounded-lg text-base font-medium text-slate-200 hover:bg-slate-800"
          >
            Overview
          </Link>
          <Link
            href="/api-reference/"
            onClick={() => setMobileMenuOpen(false)}
            className="block px-3 py-2 rounded-lg text-base font-medium text-slate-200 hover:bg-slate-800"
          >
            API Reference
          </Link>
          <Link
            href="/guides/quick-start/"
            onClick={() => setMobileMenuOpen(false)}
            className="block px-3 py-2 rounded-lg text-base font-medium text-slate-200 hover:bg-slate-800"
          >
            Quick Start Guide
          </Link>
          <Link
            href="/guides/serverless-deployment/"
            onClick={() => setMobileMenuOpen(false)}
            className="block px-3 py-2 rounded-lg text-base font-medium text-slate-200 hover:bg-slate-800"
          >
            Serverless Guide
          </Link>
          <div className="pt-4 border-t border-slate-800/80 flex items-center gap-3">
            <a
              href={navigation.external.github}
              target="_blank"
              rel="noopener noreferrer"
              className="flex-1 text-center py-2 rounded-lg text-xs font-semibold bg-slate-900 text-slate-200 border border-slate-700"
            >
              GitHub
            </a>
            <a
              href={navigation.external.npm}
              target="_blank"
              rel="noopener noreferrer"
              className="flex-1 text-center py-2 rounded-lg text-xs font-semibold bg-rose-950/40 text-rose-300 border border-rose-800/40"
            >
              NPM Package
            </a>
          </div>
        </div>
      )}
    </header>
  );
}
