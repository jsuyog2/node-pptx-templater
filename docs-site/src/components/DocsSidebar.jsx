'use client';

import React, { useState } from 'react';
import Link from 'next/link';
import { usePathname } from 'next/navigation';
import {
  BookOpen,
  Download,
  Cpu,
  Code2,
  ChevronRight,
  ChevronDown,
  Layers,
  Sparkles,
  Table,
  BarChart3,
  Image as ImageIcon,
  Split,
  ShieldCheck,
  Menu,
  X
} from 'lucide-react';
import docsNavigation from '../data/docsNavigation.json';
import apiReference from '../data/apiReference.json';

const CATEGORY_ICONS = {
  'Lifecycle & I/O': Layers,
  'Slides': Layers,
  'Text & Placeholders': Sparkles,
  'Tables & Cell Shapes': Table,
  'Charts': BarChart3,
  'Shapes & Geometry': Cpu,
  'Images & Media': ImageIcon,
  'Z-Order & Layering': Split,
  'Hyperlinks': ChevronRight,
  'Diagnostics & Validation': ShieldCheck
};

export default function DocsSidebar() {
  const pathname = usePathname();
  const [mobileOpen, setMobileOpen] = useState(false);
  const [apiCategoriesOpen, setApiCategoriesOpen] = useState(true);

  const isOverview = pathname === '/docs';
  const isInstallation = pathname === '/docs/installation';
  const isCompatibility = pathname === '/docs/compatibility';
  const isApi = pathname === '/docs/api' || pathname.startsWith('/api-reference');

  const sidebarContent = (
    <div className="h-full flex flex-col justify-between p-4 space-y-6">
      <div className="space-y-6">
        {/* Getting Started Section */}
        <div>
          <h4 className="text-[11px] font-bold uppercase tracking-wider text-slate-500 dark:text-slate-400 mb-2.5 px-2">
            Getting Started
          </h4>
          <nav className="space-y-1">
            <Link
              href="/docs"
              onClick={() => setMobileOpen(false)}
              className={`flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-medium transition-all ${isOverview
                ? 'bg-brand-500/15 text-brand-700 dark:text-brand-300 font-semibold border border-brand-500/25 shadow-sm'
                : 'text-slate-600 dark:text-slate-300 hover:text-slate-900 dark:hover:text-white hover:bg-slate-100 dark:hover:bg-slate-800/60'
                }`}
            >
              <span className="text-base">📌</span>
              <span>Overview</span>
            </Link>

            <Link
              href="/docs/installation"
              onClick={() => setMobileOpen(false)}
              className={`flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-medium transition-all ${isInstallation
                ? 'bg-brand-500/15 text-brand-700 dark:text-brand-300 font-semibold border border-brand-500/25 shadow-sm'
                : 'text-slate-600 dark:text-slate-300 hover:text-slate-900 dark:hover:text-white hover:bg-slate-100 dark:hover:bg-slate-800/60'
                }`}
            >
              <span className="text-base">🚀</span>
              <span>Installation Guides</span>
            </Link>

            <Link
              href="/docs/compatibility"
              onClick={() => setMobileOpen(false)}
              className={`flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-medium transition-all ${isCompatibility
                ? 'bg-brand-500/15 text-brand-700 dark:text-brand-300 font-semibold border border-brand-500/25 shadow-sm'
                : 'text-slate-600 dark:text-slate-300 hover:text-slate-900 dark:hover:text-white hover:bg-slate-100 dark:hover:bg-slate-800/60'
                }`}
            >
              <span className="text-base">⚡</span>
              <span>Compatibility Matrix</span>
            </Link>
          </nav>
        </div>

        {/* API Reference & Sub-Categories Section */}
        <div>
          <div className="flex items-center justify-between px-2 mb-2">
            <Link
              href="/docs/api"
              onClick={() => setMobileOpen(false)}
              className="flex items-center gap-1.5 text-[11px] font-bold uppercase tracking-wider text-slate-500 dark:text-slate-400 hover:text-brand-500"
            >
              <span className="text-sm">📚</span>
              <span>API Reference</span>
            </Link>
            <button
              onClick={() => setApiCategoriesOpen(!apiCategoriesOpen)}
              className="p-1 text-slate-400 hover:text-slate-600 dark:hover:text-slate-200"
              aria-label="Toggle categories"
            >
              {apiCategoriesOpen ? <ChevronDown className="w-3.5 h-3.5" /> : <ChevronRight className="w-3.5 h-3.5" />}
            </button>
          </div>

          <div className="mb-1">
            <Link
              href="/docs/api"
              onClick={() => setMobileOpen(false)}
              className={`w-full flex items-center justify-between px-3 py-1.5 rounded-xl text-xs transition-all ${isApi && !pathname.includes('#')
                ? 'bg-brand-500/15 text-brand-700 dark:text-brand-300 font-semibold'
                : 'text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800/60'
                }`}
            >
              <span>All Methods</span>
              <span className="text-[10px] font-mono px-1.5 py-0.5 rounded bg-slate-200 dark:bg-slate-800 text-slate-700 dark:text-slate-300">
                {Object.values(apiReference.methodsByCategory).flatMap(methods => methods).length}
              </span>
            </Link>
          </div>

          {apiCategoriesOpen && (
            <div className="space-y-0.5 pl-2 border-l border-slate-200 dark:border-slate-800 ml-2">
              {docsNavigation.sidebar[1].items.map((cat) => {
                const Icon = CATEGORY_ICONS[cat.category] || Code2;
                return (
                  <Link
                    key={cat.category}
                    href={`/docs/api?category=${encodeURIComponent(cat.category)}`}
                    onClick={() => setMobileOpen(false)}
                    className="flex items-center justify-between px-2.5 py-1.5 rounded-lg text-xs text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white hover:bg-slate-100 dark:hover:bg-slate-800/40 transition-colors"
                  >
                    <span className="truncate pr-1 flex items-center gap-1"><Icon className="size-3" />{cat.title}</span>
                    <span className="text-[10px] font-mono text-slate-400 dark:text-slate-500 shrink-0">
                      {apiReference.methodsByCategory[cat.category].length}
                    </span>
                  </Link>
                );
              })}
            </div>
          )}
        </div>
      </div>

      {/* Footer Info in Sidebar */}
      <div className="pt-4 border-t border-slate-200 dark:border-slate-800 text-[11px] text-slate-400">
        <p className="font-mono">v3.0.0-alpha.0</p>
        <p className="text-[10px] text-slate-500 mt-0.5">Zero Microsoft Office dependencies</p>
      </div>
    </div>
  );

  return (
    <>
      {/* Desktop Fixed Always-on-Top Sidebar */}
      <aside className="hidden lg:block fixed top-16 left-0 h-[calc(100vh-4rem)] w-72 z-40 overflow-y-auto glass-sidebar transition-colors">
        {sidebarContent}
      </aside>

      {/* Mobile Sidebar Floating Pill Trigger */}
      <div className="lg:hidden fixed bottom-6 right-6 z-50">
        <button
          onClick={() => setMobileOpen(!mobileOpen)}
          className="flex items-center gap-2 px-4 py-2.5 rounded-full bg-brand-600 text-white shadow-xl shadow-brand-500/30 text-xs font-semibold hover:bg-brand-500 transition-all"
          aria-label="Open documentation navigation"
        >
          {mobileOpen ? <X className="w-4 h-4" /> : <Menu className="w-4 h-4" />}
          <span>Docs Menu</span>
        </button>
      </div>

      {/* Mobile Drawer Overlay */}
      {mobileOpen && (
        <div className="lg:hidden fixed inset-0 z-40 bg-slate-950/60 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="fixed top-16 left-0 bottom-0 w-80 max-w-[85vw] glass-sidebar z-50 overflow-y-auto p-2">
            {sidebarContent}
          </div>
        </div>
      )}
    </>
  );
}
