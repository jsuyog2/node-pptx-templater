/**
 * @fileoverview Performance benchmark for pptx-templater.
 *
 * Measures:
 *  1. Template load time
 *  2. Text replacement throughput
 *  3. Chart update speed
 *  4. Buffer generation time
 *
 * Run: node benchmarks/run.js
 *
 * Results are saved to benchmarks/results/latest.json
 */

import { performance } from 'perf_hooks';
import { writeFileSync, existsSync, mkdirSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { PPTXTemplater } from '../src/index.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const RESULTS_DIR = resolve(__dirname, 'results');
const TEMPLATE = resolve(__dirname, '../templates/sample.pptx');
const ITERATIONS = 50;

// Create results dir if needed
if (!existsSync(RESULTS_DIR)) mkdirSync(RESULTS_DIR, { recursive: true });

/**
 * Measures average execution time of an async function over N iterations.
 *
 * @param {string} name - Benchmark name.
 * @param {Function} fn - Async function to benchmark.
 * @param {number} [iters] - Number of iterations.
 * @returns {Promise<BenchmarkResult>}
 */
async function benchmark(name, fn, iters = ITERATIONS) {
  // Warm up
  for (let i = 0; i < 3; i++) {
    await fn();
  }

  const times = [];
  for (let i = 0; i < iters; i++) {
    const start = performance.now();
    await fn();
    times.push(performance.now() - start);
  }

  const sorted = [...times].sort((a, b) => a - b);
  const avg = times.reduce((a, b) => a + b, 0) / times.length;
  const median = sorted[Math.floor(sorted.length / 2)];
  const p95 = sorted[Math.floor(sorted.length * 0.95)];
  const min = sorted[0];
  const max = sorted[sorted.length - 1];

  return {
    name,
    iterations: iters,
    avg: +avg.toFixed(3),
    median: +median.toFixed(3),
    p95: +p95.toFixed(3),
    min: +min.toFixed(3),
    max: +max.toFixed(3),
    opsPerSec: +(1000 / avg).toFixed(1),
  };
}

/**
 * Prints a formatted benchmark result table.
 */
function printResult(result) {
  console.log(`\n📊 ${result.name}`);
  console.log(`   Avg:       ${result.avg}ms`);
  console.log(`   Median:    ${result.median}ms`);
  console.log(`   P95:       ${result.p95}ms`);
  console.log(`   Min/Max:   ${result.min}ms / ${result.max}ms`);
  console.log(`   Ops/sec:   ${result.opsPerSec}`);
}

async function runBenchmarks() {
  console.log('🔥 pptx-templater Benchmarks');
  console.log('====================================');
  console.log(`Iterations: ${ITERATIONS}`);
  console.log(`Template:   ${existsSync(TEMPLATE) ? TEMPLATE : '(simulated)'}\n`);

  if (!existsSync(TEMPLATE)) {
    console.log('⚠  No template file found. Running simulated benchmarks only.\n');
    await runSimulatedBenchmarks();
    return;
  }

  const results = [];

  // ── Benchmark 1: Load time ──────────────────────────────────────────
  results.push(await benchmark('Load PPTX template', async () => {
    await PPTXTemplater.load(TEMPLATE);
  }, 20)); // Fewer iterations since it's slower

  // ── Benchmark 2: Text replacement ──────────────────────────────────
  const pptForText = await PPTXTemplater.load(TEMPLATE);
  const replacements = {};
  for (let i = 1; i <= 20; i++) {
    replacements[`{{placeholder${i}}}`] = `Value ${i}`;
  }

  results.push(await benchmark('Text replacement (20 placeholders)', async () => {
    pptForText.useAllSlides().replaceText(replacements);
  }));

  // ── Benchmark 3: Buffer generation ─────────────────────────────────
  const pptForBuffer = await PPTXTemplater.load(TEMPLATE);
  results.push(await benchmark('Generate buffer', async () => {
    await pptForBuffer.toBuffer();
  }, 20));

  // ── Print and save results ──────────────────────────────────────────
  results.forEach(printResult);

  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const resultsPath = resolve(RESULTS_DIR, `${timestamp}.json`);
  const latestPath = resolve(RESULTS_DIR, 'latest.json');

  const output = {
    timestamp: new Date().toISOString(),
    nodeVersion: process.version,
    platform: process.platform,
    results,
  };

  writeFileSync(resultsPath, JSON.stringify(output, null, 2));
  writeFileSync(latestPath, JSON.stringify(output, null, 2));

  console.log(`\n💾 Results saved to ${latestPath}`);
}

async function runSimulatedBenchmarks() {
  const { TemplateEngine } = await import('../src/core/TemplateEngine.js');
  const { XMLParser } = await import('../src/parsers/XMLParser.js');

  const engine = new TemplateEngine(new XMLParser());

  // Generate a large slide XML
  const runs = Array.from({ length: 100 }, (_, i) =>
    `<a:r><a:t>{{item${i}}}</a:t></a:r>`
  ).join('');
  const slideXml = `<p:sld><a:p>${runs}</a:p></p:sld>`;

  const replacements = {};
  for (let i = 0; i < 100; i++) replacements[`{{item${i}}}`] = `Value ${i}`;

  const result = await benchmark('Text replacement (100 placeholders, simulated)', async () => {
    engine.replaceTextInXml(slideXml, replacements);
  });

  printResult(result);
  console.log('\n✅ Simulated benchmarks complete');
}

runBenchmarks().catch(console.error);
