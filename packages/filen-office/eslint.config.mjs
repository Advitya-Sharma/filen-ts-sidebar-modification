import js from "@eslint/js"
import tseslint from "typescript-eslint"

// The core runs in a worker on web, inside a WebView on mobile and next to the file in Hermes, so it may
// not reach any platform: no Node, React, React Native, Expo, or the package's own React layer.
const coreImportPatterns = [
	{
		group: ["node:*", "react", "react-dom", "react-dom/*", "react-native", "react-native/*", "expo", "expo-*", "@expo/*"],
		message: "The office core is platform-free: no Node, React, React Native or Expo imports."
	},
	{
		group: ["**/react/**"],
		message: "The core may not import the React layer."
	}
]

// Nothing in the package may depend on a host app.
const appImportPatterns = [
	{
		group: ["@filen/web", "@filen/web/*", "@filen/mobile", "@filen/mobile/*", "@filen/desktop", "@filen/desktop/*"],
		message: "@filen/office is used by the apps, never the other way round."
	}
]

export default tseslint.config(
	js.configs.recommended,
	...tseslint.configs.recommended,
	{
		ignores: ["node_modules/**/*"]
	},
	{
		files: ["**/*.ts"],
		rules: {
			eqeqeq: 2,
			quotes: ["error", "double", { avoidEscape: true }],
			"no-mixed-spaces-and-tabs": 0,
			"no-duplicate-imports": "error",
			"no-extra-semi": 0,
			"@typescript-eslint/no-unused-vars": [
				"error",
				{
					argsIgnorePattern: "^_",
					varsIgnorePattern: "^_",
					caughtErrorsIgnorePattern: "^_",
					destructuredArrayIgnorePattern: "^_"
				}
			],
			"no-restricted-imports": [
				"error",
				{
					patterns: appImportPatterns
				}
			]
		}
	},
	{
		files: ["src/core/**"],
		rules: {
			"no-restricted-imports": [
				"error",
				{
					patterns: [...coreImportPatterns, ...appImportPatterns]
				}
			]
		}
	}
)
