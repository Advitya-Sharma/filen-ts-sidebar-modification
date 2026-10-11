// vite/spellcheck-assets-plugin.ts: the spellchecker's wasm module and dictionary files, as this build
// serves them (content-hashed, same origin).
declare module "virtual:spellcheck-assets" {
	export const SPELLCHECK_ASSETS: import("@filen/spellcheck").SpellcheckAssets
}
