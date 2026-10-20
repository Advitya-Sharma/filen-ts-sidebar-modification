import { readFileSync, readdirSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import * as off from "@/locales/en/previewSheet/off"
import { SpreadsheetViewer } from "@/features/spreadsheet/components/spreadsheetViewer"
import { SpreadsheetView } from "@/features/spreadsheet/next/spreadsheetView.off"

// Builds without VITE_SHEET_NEXT carry none of the office engine's viewer: vite.config.ts resolves the switch
// between the viewers to the current one, and every catalog of the new viewer's strings to an empty stub.

const ROOT = process.cwd()
const LOCALES = join(ROOT, "src/locales/en/previewSheet")

function aliases(): [RegExp, string][] {
	const config = readFileSync(join(ROOT, "vite.config.ts"), "utf8")
	const body = config.slice(config.indexOf("function sheetNextAlias("), config.indexOf("const contentSecurityPolicy"))

	return [...body.matchAll(/find: \/(.+?)\/,\s*replacement: "([^"]+)"/g)].map(match => [new RegExp(match[1] ?? ""), match[2] ?? ""])
}

function sources(directory: string, into: { path: string; text: string }[] = []): { path: string; text: string }[] {
	for (const entry of readdirSync(directory, { withFileTypes: true })) {
		const path = join(directory, entry.name)

		if (entry.isDirectory()) {
			if (entry.name !== "locales" && entry.name !== "tests") sources(path, into)
		} else if (/\.tsx?$/.test(entry.name)) into.push({ path, text: readFileSync(path, "utf8") })
	}

	return into
}

describe("the office engine's spreadsheet viewer left out of builds without it", () => {
	it("resolves the viewer switch to the current viewer itself", () => {
		const alias = aliases().find(([pattern]) => pattern.test("@/features/spreadsheet/next/spreadsheetView"))

		expect(alias?.[1]).toBe("/src/features/spreadsheet/next/spreadsheetView.off.ts")
		expect(SpreadsheetView).toBe(SpreadsheetViewer)
	})

	it("resolves every string catalog of the new viewer to the empty stub, which names each catalog", async () => {
		const found = aliases()
		const files = readdirSync(LOCALES).filter(file => file.endsWith(".ts") && file !== "off.ts")

		expect(files.length).toBeGreaterThan(0)

		for (const file of files) {
			const specifier = `@/locales/en/previewSheet/${file.replace(/\.ts$/, "")}`

			expect(found.find(([pattern]) => pattern.test(specifier))?.[1], specifier).toBe("/src/locales/en/previewSheet/off.ts")

			const catalog = (await import(/* @vite-ignore */ join(LOCALES, file))) as Record<string, object>

			for (const name of Object.keys(catalog)) {
				expect((off as Record<string, object>)[name], `${file} ${name}`).toEqual({})
			}
		}
	})

	it("reaches the new viewer only through the switch, so nothing else carries it into a build", () => {
		const outside = sources(join(ROOT, "src"))
			.filter(({ path }) => !path.includes(join("features", "spreadsheet", "next")))
			.filter(({ text }) => /from "@\/features\/spreadsheet\/next\/(?!spreadsheetView")/.test(text))
			.map(({ path }) => path)

		expect(outside).toEqual([])
	})

	it("ships no string of the new viewer the app never names", () => {
		const code = sources(join(ROOT, "src"))
			.map(({ text }) => text)
			.join("\n")
		const unused = readdirSync(LOCALES)
			.filter(file => file !== "off.ts")
			.flatMap(file =>
				[...readFileSync(join(LOCALES, file), "utf8").matchAll(/^\t(previewSheet[A-Za-z0-9]+?)(?:_one|_other)?:/gm)].map(
					match => match[1] ?? ""
				)
			)
			.filter(key => !code.includes(key))

		expect(unused).toEqual([])
	})
})
