import js from '@eslint/js';

export default [
  js.configs.recommended,
  {
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'module',
      globals: {
        process: 'readonly',
        Buffer: 'readonly',
        console: 'readonly',
        URL: 'readonly',
        fetch: 'readonly',
        setTimeout: 'readonly',
        clearTimeout: 'readonly',
        setInterval: 'readonly',
        clearInterval: 'readonly',
      },
    },
    files: ["**/*.js"],
    rules: {
      // Error prevention
      'no-unused-vars': ['warn', { argsIgnorePattern: '^_', varsIgnorePattern: '^_' }],
      'no-console': 'off', // Allow console in this library
      'no-var': 'error',
      'prefer-const': 'error',

      // Code style
      'no-trailing-spaces': 'error',
      'eol-last': ['error', 'always'],
      'no-multiple-empty-lines': ['error', { max: 2, maxEOF: 1 }],

      // ES Module best practices
      'no-duplicate-imports': 'error',

      // Disable checks on pre-existing legacy issues
      'no-empty': 'off',
      'no-useless-escape': 'off',
      'no-constant-condition': 'off',
      'no-control-regex': 'off',
    },
    ignores: [
      'node_modules/**',
      'dist/**',
      'coverage/**',
      'benchmarks/results/**',
    ],
  },
];
