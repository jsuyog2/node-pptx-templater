'use client';

import React, { useState } from 'react';
import Link from 'next/link';
import { usePathname } from 'next/navigation';
import { Layers, BookOpen, Sparkles, Code2, Github, ExternalLink, Menu, X, Star } from 'lucide-react';
import ThemeToggle from './ThemeToggle';
import siteMeta from '../data/siteMeta.json';

export default function Header() {
  const pathname = usePathname();
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);

  const navLinks = [
    { title: 'Home', href: '/', active: pathname === '/' },
    {
      title: 'Docs',
      href: '/docs',
      active: pathname.startsWith('/docs') && pathname !== '/docs/api',
      icon: BookOpen
    },
    {
      title: 'Features',
      href: '/features',
      active: pathname === '/features',
      icon: Sparkles
    },
    // {
    //   title: 'API Reference',
    //   href: '/docs/api',
    //   active: pathname === '/docs/api' || pathname.startsWith('/api-reference'),
    //   icon: Code2
    // }
  ];

  return (
    <header className="fixed top-0 inset-x-0 z-50 h-16 glass-header transition-colors">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-full">
        <div className="flex items-center justify-between h-full">
          {/* Brand Logo & Name */}
          <Link href="/" className="flex items-center gap-3 group">
            <div className="w-9 h-9 rounded-xl bg-gradient-to-tr from-brand-500 via-sky-400 to-indigo-500 p-0.5 shadow-md shadow-brand-500/20 group-hover:scale-105 transition-transform duration-200">
              <div className="w-full h-full bg-white dark:bg-slate-950 rounded-[10px] flex items-center justify-center">
                <Layers className="w-5 h-5 text-brand-600 dark:text-brand-400 group-hover:text-brand-500 transition-colors" />
              </div>
            </div>
            <div className="flex flex-col">
              <div className="flex items-center gap-2">
                <span className="font-bold text-base sm:text-lg tracking-tight text-slate-900 dark:text-white group-hover:text-brand-600 dark:group-hover:text-brand-300 transition-colors">
                  node-pptx-templater
                </span>
                <span className="hidden sm:inline-block text-[11px] font-semibold px-2 py-0.5 rounded-full bg-brand-500/10 dark:bg-brand-500/15 text-brand-600 dark:text-brand-400 border border-brand-500/20 dark:border-brand-500/30">
                  {siteMeta.version}
                </span>
              </div>
              <span className="text-[10px] text-slate-500 dark:text-slate-400 font-mono -mt-0.5 hidden sm:block">
                Pure JS OpenXML Template Engine
              </span>
            </div>
          </Link>

          {/* Desktop Navigation Tabs */}
          <nav className="hidden md:flex items-center gap-1.5">
            {navLinks.map((link) => {
              const Icon = link.icon;
              return (
                <Link
                  key={link.title}
                  href={link.href}
                  className={`flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl text-sm font-medium transition-all duration-150 ${link.active
                      ? 'bg-brand-500/15 text-brand-700 dark:text-brand-300 font-semibold border border-brand-500/25 shadow-sm'
                      : 'text-slate-600 dark:text-slate-300 hover:text-slate-900 dark:hover:text-white hover:bg-slate-100 dark:hover:bg-slate-800/60'
                    }`}
                >
                  {Icon && <Icon className={`w-4 h-4 ${link.active ? 'text-brand-600 dark:text-brand-400' : 'text-slate-400'}`} />}
                  <span>{link.title}</span>
                </Link>
              );
            })}
          </nav>

          {/* Right Action Icons (ThemeToggle + GitHub Repo) */}
          <div className="hidden sm:flex items-center gap-2.5">
            <ThemeToggle />

            <a
              href="https://github.com/jsuyog2/node-pptx-templater"
              target="_blank"
              rel="noopener noreferrer"
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-semibold border border-slate-300/80 dark:border-slate-800 bg-white/70 dark:bg-slate-900/80 hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200 transition-all shadow-sm group"
            >
              <Github className="w-4 h-4" />
              <span>GitHub</span>
              <span className="flex items-center gap-0.5 text-[10px] text-amber-500 font-mono font-normal ml-0.5">
                <Star className="w-3 h-3 fill-amber-400 text-amber-400" />
              </span>
            </a>

            <a
              href="https://www.npmjs.com/package/node-pptx-templater"
              target="_blank"
              rel="noopener noreferrer"
              className="hidden lg:flex items-center gap-1.5 px-2.5 py-1.5 rounded-xl text-xs font-semibold border border-rose-200 dark:border-rose-900/40 bg-rose-50 dark:bg-rose-950/40 hover:bg-rose-100 dark:hover:bg-rose-900/50 text-rose-700 dark:text-rose-300 transition-all shadow-sm"
            >
              <span className="font-mono font-bold text-[11px]">npm</span>
              <ExternalLink className="w-3 h-3 text-rose-500" />
            </a>
          </div>

          {/* Mobile Actions & Menu Toggle */}
          <div className="flex md:hidden items-center gap-2">
            <ThemeToggle />
            <button
              onClick={() => setMobileMenuOpen(!mobileMenuOpen)}
              className="p-2 rounded-xl text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white hover:bg-slate-100 dark:hover:bg-slate-800"
              aria-label="Toggle navigation menu"
            >
              {mobileMenuOpen ? <X className="w-6 h-6" /> : <Menu className="w-6 h-6" />}
            </button>
          </div>
        </div>
      </div>

      {/* Mobile Drawer */}
      {mobileMenuOpen && (
        <div className="md:hidden glass-header border-b px-4 pt-3 pb-6 space-y-2 shadow-xl animate-in slide-in-from-top-2 duration-200">
          {navLinks.map((link) => {
            const Icon = link.icon;
            return (
              <Link
                key={link.title}
                href={link.href}
                onClick={() => setMobileMenuOpen(false)}
                className={`flex items-center gap-2.5 px-3 py-2 rounded-xl text-sm font-medium ${link.active
                    ? 'bg-brand-500/15 text-brand-700 dark:text-brand-300 font-semibold'
                    : 'text-slate-700 dark:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800'
                  }`}
              >
                {Icon && <Icon className="w-4 h-4 text-brand-500" />}
                <span>{link.title}</span>
              </Link>
            );
          })}

          <div className="pt-3 border-t border-slate-200 dark:border-slate-800 flex items-center gap-2">
            <a
              href="https://github.com/jsuyog2/node-pptx-templater"
              target="_blank"
              rel="noopener noreferrer"
              className="flex-1 text-center py-2 rounded-xl text-xs font-semibold bg-slate-100 dark:bg-slate-800 text-slate-800 dark:text-slate-200 border border-slate-300 dark:border-slate-700 flex items-center justify-center gap-1.5"
            >
              <Github className="w-3.5 h-3.5" />
              <span>GitHub</span>
            </a>
            <a
              href="https://www.npmjs.com/package/node-pptx-templater"
              target="_blank"
              rel="noopener noreferrer"
              className="flex-1 text-center py-2 rounded-xl text-xs font-semibold bg-rose-50 dark:bg-rose-950/40 text-rose-700 dark:text-rose-300 border border-rose-200 dark:border-rose-800/40"
            >
              npm Package
            </a>
          </div>
        </div>
      )}
    </header>
  );
}
