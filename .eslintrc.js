require('@rushstack/eslint-config/patch/modern-module-resolution');

module.exports = {
  parser: require.resolve('@typescript-eslint/parser'),
  extends: [
    '@microsoft/eslint-config-spfx/lib/profiles/react'
  ],
  parserOptions: {
    tsconfigRootDir: __dirname,
    project: './tsconfig.json',
    ecmaVersion: 2020,
    sourceType: 'module',
    ecmaFeatures: {
      jsx: true
    }
  },
  rules: {
    '@typescript-eslint/no-explicit-any': 'off',
    '@typescript-eslint/explicit-function-return-type': 'off',
    '@typescript-eslint/no-unused-vars': 'off',
    'react/no-unescaped-entities': 'off',
    '@typescript-eslint/no-unused-expressions': 'off',
    '@rushstack/security/no-unsafe-regexp': 'off',
    'no-extra-boolean-cast': 'off',
    'prefer-const': 'off',
    'no-empty': 'off',
    'no-useless-escape': 'off',
    'react/self-closing-comp': 'off'
  }
};
