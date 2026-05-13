/*
 * ESLint configuration
 *
 * Goal of this file: move from an empty rule set to a real baseline without
 * blocking CI on pre-existing debt. Legitimate-for-this-codebase rules are
 * disabled (e.g. `no-control-regex` is off because security-validator
 * intentionally matches control characters). Pre-existing stylistic debt is
 * downgraded to `warn` so it shows up without breaking CI; tighten to
 * `error` in a follow-up once cleaned.
 */
module.exports = {
  root: true,
  env: {
    es2022: true,
    node: true,
  },
  parser: '@typescript-eslint/parser',
  parserOptions: {
    ecmaVersion: 'latest',
    sourceType: 'module',
  },
  plugins: ['@typescript-eslint'],
  extends: [
    'eslint:recommended',
    'plugin:@typescript-eslint/recommended',
  ],
  ignorePatterns: ['build/', 'node_modules/'],
  rules: {
    // Security validators legitimately match control characters in regexes.
    'no-control-regex': 'off',

    // Pre-existing debt: track as warning until cleaned up.
    '@typescript-eslint/no-explicit-any': 'warn',
    '@typescript-eslint/no-unused-vars': [
      'warn',
      { argsIgnorePattern: '^_', varsIgnorePattern: '^_', caughtErrorsIgnorePattern: '^_' },
    ],
    '@typescript-eslint/ban-types': 'warn',
    'no-unused-vars': 'off',
    'no-useless-escape': 'warn',
    'no-case-declarations': 'warn',

    // Empty function bodies appear in mocks and interface placeholders.
    '@typescript-eslint/no-empty-function': 'off',
    // Scripts use CommonJS requires; TS checker handles type imports.
    '@typescript-eslint/no-var-requires': 'off',
  },
  overrides: [
    {
      files: ['scripts/**/*.mjs', '*.cjs'],
      parser: 'espree',
    },
  ],
};
