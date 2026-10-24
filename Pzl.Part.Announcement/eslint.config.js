const tsParser = require('@typescript-eslint/parser')
const tsPlugin = require('@typescript-eslint/eslint-plugin')
const react = require('eslint-plugin-react')
const prettierRecommended = require('eslint-plugin-prettier/recommended')
const unusedImports = require('eslint-plugin-unused-imports')

module.exports = [
  {
    ignores: [
      'node_modules/**',
      'lib/**',
      'lib-commonjs/**',
      'lib-esm/**',
      'lib-dts/**',
      'dist/**',
      'release/**',
      'temp/**',
      'sharepoint/**',
      'src/**/*.d.ts',
      'src/**/*.scss.ts',
      'eslint.config.js'
    ]
  },
  {
    settings: {
      react: {
        version: 'detect'
      }
    }
  },
  react.configs.flat.recommended,
  ...tsPlugin.configs['flat/recommended'],
  prettierRecommended,
  {
    files: ['**/*.ts', '**/*.tsx'],
    plugins: {
      'unused-imports': unusedImports
    },
    languageOptions: {
      parser: tsParser,
      ecmaVersion: 2018,
      sourceType: 'module',
      parserOptions: {
        ecmaFeatures: {
          jsx: true
        }
      }
    },
    rules: {
      '@typescript-eslint/no-explicit-any': 0,
      '@typescript-eslint/no-unused-vars': 1,
      '@typescript-eslint/explicit-function-return-type': 0,
      '@typescript-eslint/explicit-module-boundary-types': 0,
      '@typescript-eslint/no-inferrable-types': 0,
      '@typescript-eslint/no-require-imports': 0,
      'react/prop-types': 0,
      'react/display-name': 0,
      'no-compare-neg-zero': 1,
      'no-unused-vars': 0,
      'no-console': 1,
      'default-case': 0,
      eqeqeq: 1,
      'max-classes-per-file': 1,
      'jsx-quotes': [0, 'prefer-single'],
      quotes: ['error', 'single'],
      yoda: 2,
      'require-await': 1,
      semi: ['error', 'never', {}],
      'unused-imports/no-unused-imports': 0
    }
  }
]
