#!/usr/bin/env node
/**
 * @fileoverview CLI entry point for node-pptx-templater.
 *
 * Provides command-line access to the template engine's core features.
 *
 * Usage:
 *   node-pptx-templater build template.pptx output.pptx [options]
 *   node-pptx-templater validate template.pptx
 *   node-pptx-templater inspect template.pptx
 *   node-pptx-templater extract template.pptx --slide 1 --out ./slide1.xml
 *   node-pptx-templater debug template.pptx
 *
 * Install globally:
 *   npm install -g node-pptx-templater
 *
 * Then run:
 *   node-pptx-templater --help
 */

import { Command } from 'commander';
import chalk from 'chalk';
import ora from 'ora';
import { readFileSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { PPTXTemplater } from '../index.js';
import { buildCommand } from './commands/build.js';
import { validateCommand } from './commands/validate.js';
import { inspectCommand } from './commands/inspect.js';
import { extractCommand } from './commands/extract.js';
import { debugCommand } from './commands/debug.js';

const __dirname = dirname(fileURLToPath(import.meta.url));

// Read version from package.json
const pkg = JSON.parse(
  readFileSync(resolve(__dirname, '../../package.json'), 'utf-8')
);

/**
 * CLI banner displayed on startup.
 */
function printBanner() {
  console.log(chalk.bold.cyan(`
╔═══════════════════════════════════════════════╗
║  node-pptx-templater v${pkg.version.padEnd(17)}║
║  Low-level OpenXML PowerPoint template engine ║
╚═══════════════════════════════════════════════╝
`));
}

const program = new Command();

program
  .name('node-pptx-templater')
  .description('Low-level PowerPoint OpenXML template engine for Node.js')
  .version(pkg.version, '-v, --version', 'Display version number')
  .addHelpText('before', chalk.bold.cyan('\nnode-pptx-templater — PowerPoint XML manipulation engine\n'));

// ─── build command ─────────────────────────────────────────────────────────
program
  .command('build <template> <output>')
  .description('Build a PPTX from a template with data injected from a JSON file')
  .option('-d, --data <file>', 'JSON file with template data (text replacements, etc.)')
  .option('-s, --slide <numbers>', 'Comma-separated slide numbers to include (e.g., 1,3,5)')
  .option('--no-banner', 'Suppress the banner')
  .action(async (template, output, opts) => {
    if (!opts.noBanner) printBanner();
    await buildCommand(template, output, opts);
  });

// ─── validate command ───────────────────────────────────────────────────────
program
  .command('validate <file>')
  .description('Validate the structure of a PPTX file')
  .option('--strict', 'Exit with error code on warnings too')
  .action(async (file, opts) => {
    printBanner();
    await validateCommand(file, opts);
  });

// ─── inspect command ────────────────────────────────────────────────────────
program
  .command('inspect <file>')
  .description('Inspect the internal structure of a PPTX file')
  .option('--slides', 'Show slide details')
  .option('--charts', 'Show chart details')
  .option('--tables', 'Show table details')
  .option('--media', 'Show embedded media files')
  .option('--rels', 'Show relationship tree')
  .option('--all', 'Show everything')
  .action(async (file, opts) => {
    printBanner();
    await inspectCommand(file, opts);
  });

// ─── extract command ────────────────────────────────────────────────────────
program
  .command('extract <file>')
  .description('Extract specific parts from a PPTX file')
  .option('-s, --slide <number>', 'Slide number to extract XML from')
  .option('-o, --out <path>', 'Output file path (default: stdout)')
  .option('--chart <name>', 'Extract chart XML by name')
  .option('--rels', 'Extract relationship files')
  .action(async (file, opts) => {
    await extractCommand(file, opts);
  });

// ─── debug command ──────────────────────────────────────────────────────────
program
  .command('debug <file>')
  .description('Debug a potentially corrupted PPTX structure')
  .option('--fix', 'Attempt automatic repairs')
  .option('-o, --out <path>', 'Output repaired file path')
  .action(async (file, opts) => {
    printBanner();
    await debugCommand(file, opts);
  });

// ─── Global error handling ──────────────────────────────────────────────────
program.exitOverride();

try {
  await program.parseAsync(process.argv);
} catch (err) {
  if (err.code === 'commander.helpDisplayed' || err.code === 'commander.version') {
    process.exit(0);
  }
  console.error(chalk.red(`\n✗ Error: ${err.message}\n`));
  process.exit(1);
}
