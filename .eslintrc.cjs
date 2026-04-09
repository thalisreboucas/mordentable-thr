module.exports = {
    root: true,
    extends: ["plugin:powerbi-visuals/recommended"],
    parserOptions: {
        project: ["./tsconfig.json"],
        tsconfigRootDir: __dirname,
    },
    ignorePatterns: ["node_modules/**", "dist/**", ".vscode/**", ".tmp/**"],
};
