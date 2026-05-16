/**
 * @fileoverview Vitest configuration for pptx-templater.
 */

import { defineConfig } from 'vitest/config';

export default defineConfig({
  esbuild: {
    // Tell esbuild to target modern Node.js so private class fields (#field) are supported
    target: 'node18',
  },
  test: {
    // Use Node.js environment
    environment: 'node',
    // Use the esbuild transform which supports private class fields natively
    pool: 'forks',

    // Test file patterns
    include: [
      'tests/unit/**/*.test.js',
      'tests/integration/**/*.test.js',
      'tests/snapshot/**/*.test.js',
    ],

    // Exclude benchmarks and examples
    exclude: [
      'benchmarks/**',
      'examples/**',
      'node_modules/**',
    ],

    // Coverage configuration
    coverage: {
      provider: 'v8',
      reporter: ['text', 'lcov', 'html'],
      include: ['src/**/*.js'],
      exclude: [
        'src/cli/**',
        'src/templates/blankPptx.js',
      ],
      // Thresholds for CI enforcement
      thresholds: {
        lines: 70,
        functions: 70,
        branches: 65,
        statements: 70,
      },
    },

    // Timeout for integration tests (PPTX processing can take a bit)
    testTimeout: 30000,

    // Report format
    reporters: ['verbose'],

    // Global setup
    globalSetup: [],

    // Snapshot settings
    snapshotOptions: {
      snapshotFormat: {
        printBasicPrototype: false,
      },
    },
  },
});
