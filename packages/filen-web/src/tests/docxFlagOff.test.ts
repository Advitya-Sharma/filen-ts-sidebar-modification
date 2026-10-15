import { readFileSync, readdirSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import * as off from "@/locales/en/previewDocx/off"
import { preview } from "@/locales/en/preview"

// Builds without DOCX editing (VITE_DOCX_EDITING unset) carry none of the editor: vite.config.ts resolves the
// editor module and every catalog of its strings (locales/en/previewDocx/*) to stubs, and nothing outside the
// editor pulls a catalog in for a type.

const ROOT = process.cwd()
const LOCALES = join(ROOT, "src/locales/en/previewDocx")

// The aliases docxEditingAlias gives a build without the flag: [pattern, replacement].
function aliases(): [RegExp, string][] {
	const config = readFileSync(join(ROOT, "vite.config.ts"), "utf8")
	const body = config.slice(config.indexOf("function docxEditingAlias("), config.indexOf("const contentSecurityPolicy"))

	return [...body.matchAll(/find: \/(.+?)\/,\s*replacement: "([^"]+)"/g)].map(match => [new RegExp(match[1] ?? ""), match[2] ?? ""])
}

describe("DOCX editing left out of builds without it", () => {
	it("resolves every editor string catalog to the empty stub, which names each catalog", async () => {
		const found = aliases()
		const files = readdirSync(LOCALES).filter(file => file.endsWith(".ts") && file !== "off.ts")

		expect(files.length).toBeGreaterThan(0)

		for (const file of files) {
			const specifier = `@/locales/en/previewDocx/${file.replace(/\.ts$/, "")}`
			const alias = found.find(([pattern]) => pattern.test(specifier))

			expect(alias?.[1], specifier).toBe("/src/locales/en/previewDocx/off.ts")

			const catalog = (await import(/* @vite-ignore */ join(LOCALES, file))) as Record<string, object>

			for (const name of Object.keys(catalog)) {
				expect((off as Record<string, object>)[name], `${file} ${name}`).toEqual({})
			}
		}

		expect(found.some(([pattern]) => pattern.test("@/features/preview/components/docxEditor"))).toBe(true)
		// A catalog named in camel case is one too.
		expect(found.find(([pattern]) => pattern.test("@/locales/en/previewDocx/pictureMenu"))?.[1]).toBe(
			"/src/locales/en/previewDocx/off.ts"
		)
	})

	it("keeps the editor's own strings in its catalog, not in the viewer's preview.ts", () => {
		const source = readFileSync(join(ROOT, "src/locales/en/preview.ts"), "utf8")

		for (const key of ["previewDocxEditToolbar", "previewDocxSaveRefused", "previewDocxCommandFailed", "previewDocxClipboardBlocked"]) {
			expect(source, key).not.toMatch(new RegExp(`^\\s+${key}:`, "m"))
			expect(preview).toHaveProperty(key)
		}
	})

	it("ships no editor string the app never names", () => {
		const sources: string[] = []
		const walk = (directory: string): void => {
			for (const entry of readdirSync(directory, { withFileTypes: true })) {
				const path = join(directory, entry.name)

				if (entry.isDirectory()) {
					if (entry.name !== "locales" && entry.name !== "tests") walk(path)
				} else if (/\.tsx?$/.test(entry.name)) sources.push(readFileSync(path, "utf8"))
			}
		}

		walk(join(ROOT, "src"))

		const code = sources.join("\n")
		const unused = readdirSync(LOCALES)
			.filter(file => file !== "off.ts")
			.flatMap(file =>
				[...readFileSync(join(LOCALES, file), "utf8").matchAll(/^\t(previewDocx[A-Za-z0-9]+?)(?:_one|_other)?:/gm)].map(
					match => match[1] ?? ""
				)
			)
			.filter(key => !code.includes(key))

		expect(unused).toEqual([])
	})

	it("imports the i18n catalogs into the app shell's logic (and the workers using it) for types only", () => {
		const source = readFileSync(join(ROOT, "src/features/shell/lib/appShell.logic.ts"), "utf8")

		expect(source).not.toMatch(/^import (?!type )[^\n]*from "@\/lib\/i18n"/m)
	})
})
