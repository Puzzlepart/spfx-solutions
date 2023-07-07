require('@rushstack/eslint-config/patch/modern-module-resolution');
module.exports = {
  extends: ['@microsoft/eslint-config-spfx/lib/profiles/react'],
  parserOptions: { tsconfigRootDir: __dirname },
  rules: {
    "@typescript-eslint/no-inferrable-types": "off",
    "@typescript-eslint/no-explicit-any": "off",
    "@typescript-eslint/no-unused-vars": "off",
    "react/no-direct-mutation-state": "off",
    "react/jsx-no-target-blank": "off",
    "react/self-closing-comp": "off",
    "react/jsx-key": "off",

    "react/no-deprecated": "off" // NOTE componentWillMount
  }
};
