import { readFileSync, readdirSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"

// The core runs in a worker, a WebView and Hermes: it type-checks against ESNext alone (tsconfig.json is the
// gate), and it never reaches a window or document through globalThis, which the types cannot see.

const PACKAGE = new URL("../../../", import.meta.url).pathname
const CORE = join(PACKAGE, "src/core")
const DOM_GLOBAL =
	/globalThis\s*(?:\.\s*|\[\s*["'`])(?:window|document|self|navigator|location|localStorage)\b|Reflect\.get\(\s*globalThis\s*,\s*["'`](?:window|document|self|navigator|location|localStorage)["'`]/

function sources(dir: string): string[] {
	return readdirSync(dir, { withFileTypes: true, recursive: true })
		.filter(entry => entry.isFile() && entry.name.endsWith(".ts"))
		.map(entry => join(entry.parentPath, entry.name))
}

describe("core stays platform-free", () => {
	it("type-checks against ESNext without DOM or Node typings", () => {
		const config = JSON.parse(readFileSync(join(PACKAGE, "tsconfig.json"), "utf8")) as {
			readonly exclude: readonly string[]
			readonly compilerOptions: { readonly lib: readonly string[]; readonly types: readonly string[] }
		}

		expect(config.compilerOptions.lib).toEqual(["ESNext"])
		expect(config.compilerOptions.types).toEqual([])
		expect(config.exclude).not.toContain("src/core")
	})

	it("never reaches a window or document through globalThis", () => {
		const files = sources(CORE)
		const leaks = files.filter(file => DOM_GLOBAL.test(readFileSync(file, "utf8")))

		expect(files.length).toBeGreaterThan(100)
		expect(leaks).toEqual([])
	})
})
