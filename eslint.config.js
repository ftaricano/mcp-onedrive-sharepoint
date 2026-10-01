/*
 * ESLint flat configuration (ESLint 9+; migrated from .eslintrc.cjs/.eslintignore).
 *
 * Goal of this file: move from an empty rule set to a real baseline without
 * blocking CI on pre-existing debt. Legitimate-for-this-codebase rules are
 * disabled (e.g. `no-control-regex` is off because security-validator
 * intentionally matches control characters). Pre-existing stylistic debt is
 * downgraded to `warn` so it shows up without breaking CI; tighten to
 * `error` in a follow-up once cleaned.
 */
import js from "@eslint/js";
import tsPlugin from "@typescript-eslint/eslint-plugin";
import globals from "globals";

export default [
  { ignores: ["build/", "node_modules/"] },
  js.configs.recommended,
  ...tsPlugin.configs["flat/recommended"],
  {
    languageOptions: {
      ecmaVersion: "latest",
      sourceType: "module",
      globals: { ...globals.es2022, ...globals.node },
    },
    rules: {
      // Security validators legitimately match control characters in regexes.
      "no-control-regex": "off",

      // Pre-existing debt: track as warning until cleaned up.
      "@typescript-eslint/no-explicit-any": "warn",
      "@typescript-eslint/no-unused-vars": [
        "warn",
        {
          argsIgnorePattern: "^_",
          varsIgnorePattern: "^_",
          caughtErrorsIgnorePattern: "^_",
          // typescript-eslint v8 changed the default to "all"; keep the v6 behaviour.
          caughtErrors: "none",
        },
      ],
      // `ban-types` was split into these three rules in typescript-eslint v8.
      "@typescript-eslint/no-empty-object-type": "warn",
      "@typescript-eslint/no-unsafe-function-type": "warn",
      "@typescript-eslint/no-wrapper-object-types": "warn",
      "no-unused-vars": "off",
      "no-useless-escape": "warn",
      "no-case-declarations": "warn",
      // New in eslint:recommended (ESLint 10): rethrown errors without `cause`.
      "preserve-caught-error": "warn",

      // Empty function bodies appear in mocks and interface placeholders.
      "@typescript-eslint/no-empty-function": "off",
      // Scripts use CommonJS requires; TS checker handles type imports.
      // (`no-var-requires` was replaced by `no-require-imports` in v8.)
      "@typescript-eslint/no-require-imports": "off",
    },
  },
];
