import eslintjs from '@eslint/js';
import tseslint from 'typescript-eslint';
import react from 'eslint-plugin-react';
import globals from 'globals';
export default [
 {ignores:['**/generated/**','**/dist/**','**/node_modules/**','mockup/**']},
 eslintjs.configs.recommended,
 ...tseslint.configs.recommended,
 {files:['**/*.ts','**/*.tsx'],plugins:{react},languageOptions:{globals:{...globals.browser,...globals.node,ComponentFramework:'readonly',...globals.jest}},settings:{react:{version:'16.14'}},rules:{'no-console':'error','@typescript-eslint/no-explicit-any':'error','@typescript-eslint/no-unused-vars':['error',{argsIgnorePattern:'^_',varsIgnorePattern:'^_'}],'react/jsx-key':'error'}}
];
