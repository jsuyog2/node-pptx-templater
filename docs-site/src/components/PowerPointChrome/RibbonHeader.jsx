'use client';

import React, { useState } from 'react';
import Link from 'next/link';
import { usePathname } from 'next/navigation';
import {
  FileText,
  Download,
  Terminal,
  Cpu,
  Code2,
  Copy,
  Check,
  ExternalLink,
  Github,
  Play,
  Layers,
  Sparkles
} from 'lucide-react';

export default function RibbonHeader() {
  const pathname = usePathname();
  const [copied, setCopied] = useState(false);

  const tabs = [
    { title: 'Overview', href: '/docs', icon: FileText },
    { title: 'Installation', href: '/docs/installation', icon: Terminal },
    { title: 'Capabilities', href: '/docs/capabilities', icon: Cpu },
    { title: 'API Reference', href: '/docs/api-reference', icon: Code2 },
  ];

  const handleQuickCopy = async () => {
    try {
      await navigator.clipboard.writeText('npm install node-pptx-templater');
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      // ignore
    }
  };

  const isActive = (href) => {
    if (href === '/docs') {
      return pathname === '/docs' || pathname === '/docs/';
    }
    return pathname.startsWith(href);
  };

  return (
    <div className="ppt-chrome-surface border-b border-[var(--ppt-chrome-border)] flex flex-col justify-between px-2 sm:px-4 py-1 select-none z-40 sticky top-[36px]">
      {/* Top Row: File Tab + Primary Navigation Tabs + Actions */}
      <div className="flex items-center justify-between gap-2 overflow-x-auto no-scrollbar">
        {/* Left: File Button & Primary Tabs */}
        <div className="flex items-center gap-1 sm:gap-2">
          {/* Persistent File Tab */}
          <Link
            href="/"
            className="flex items-center gap-1 px-3 py-1 rounded-sm text-xs font-semibold bg-[#D24726] text-white hover:bg-[#b83b1e] transition-colors shadow-sm"
          >
            <span>Home</span>
          </Link>

          {/* Navigation Tabs */}
          <nav className="flex items-center gap-0.5">
            {tabs.map((tab) => {
              const active = isActive(tab.href);
              const Icon = tab.icon;
              return (
                <Link
                  key={tab.href}
                  href={tab.href}
                  className={`flex items-center gap-1.5 px-3 py-1 text-xs font-medium rounded-t-sm transition-all relative ${active
                    ? 'text-[var(--ppt-text-primary)] font-semibold border-b-2 border-[#D24726]'
                    : 'text-[var(--ppt-text-secondary)] hover:text-[var(--ppt-text-primary)] hover:bg-black/5 dark:hover:bg-white/5'
                    }`}
                >
                  <Icon className={`w-3.5 h-3.5 ${active ? 'text-[#D24726]' : 'opacity-70'}`} />
                  <span>{tab.title}</span>
                </Link>
              );
            })}
          </nav>
        </div>

        {/* Right: Ribbon Quick Tools */}
        <div className="flex items-center gap-1.5 sm:gap-2 text-xs">
          {/* Quick Install Pill */}
          <button
            type="button"
            onClick={handleQuickCopy}
            className="hidden md:flex items-center gap-1.5 px-2.5 py-1 rounded text-[11px] font-mono bg-black/5 dark:bg-white/5 border border-[var(--ppt-chrome-border)] hover:border-[#D24726] text-[var(--ppt-text-primary)] transition-all"
            title="Click to copy: npm install node-pptx-templater"
          >
            {copied ? (
              <>
                <Check className="w-3 h-3 text-emerald-500" />
                <span className="text-emerald-500 font-sans font-medium">Copied!</span>
              </>
            ) : (
              <>
                <span className="text-[#D24726] font-bold">$</span>
                <span>npm i node-pptx-templater</span>
                <Copy className="w-3 h-3 text-[var(--ppt-text-secondary)] ml-0.5" />
              </>
            )}
          </button>

          {/* GitHub Repo */}
          <a
            href="https://github.com/jsuyog2/node-pptx-templater"
            target="_blank"
            rel="noopener noreferrer"
            className="flex items-center gap-1 px-2 py-1 rounded text-xs font-medium text-[var(--ppt-text-secondary)] hover:text-[var(--ppt-text-primary)] hover:bg-black/5 dark:hover:bg-white/5 transition-colors"
            title="GitHub Repository"
          >
            <Github className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">GitHub</span>
          </a>

          {/* npm Badge Link */}
          <a
            href="https://www.npmjs.com/package/node-pptx-templater"
            target="_blank"
            rel="noopener noreferrer"
            className="flex items-center gap-1 px-2 py-1 rounded text-xs font-medium text-[var(--ppt-text-secondary)] hover:text-[var(--ppt-text-primary)] hover:bg-black/5 dark:hover:bg-white/5 transition-colors"
            title="View on npm"
          >
            <span className="font-bold text-[#D24726] text-[11px]">npm</span>
            <ExternalLink className="w-3 h-3 opacity-60" />
          </a>

          <div className="flex items-center gap-1 text-[10px]">
            <span className="hidden sm:inline opacity-70">Engine:</span>
            <span className="font-semibold text-[var(--ppt-text-primary)]">v3.0.0-beta.0</span>
          </div>
        </div>
      </div>
    </div>
  );
}
