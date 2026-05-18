import powerbiVisualsConfigs from "eslint-plugin-powerbi-visuals";
import path from "node:path";
import { fileURLToPath } from "node:url";
import typescriptEslintParser from "@typescript-eslint/parser";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export default [
    powerbiVisualsConfigs.configs.recommended,
    {
        files: ["src/**/*.ts"],
        languageOptions: {
            parser: typescriptEslintParser,
            ecmaVersion: 2020,
            sourceType: "module",
            parserOptions: {
                tsconfigRootDir: __dirname,
                project: "./tsconfig.json",
            },
        },
    },
    {
        ignores: ["node_modules/**", "dist/**", ".vscode/**", ".tmp/**"],
    },
];