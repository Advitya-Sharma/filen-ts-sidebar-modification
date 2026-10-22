import { readFileSync, readdirSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"

// Strings the spreadsheet viewer shares with DOCX editing live in a catalogue of their own (locales/en/previewOffice/*).

const ROOT = process.cwd()

function sources(directory: string, into: { path: string; text: string }[] = []): { path: string; text: string }[] {
	for (const entry of readdirSync(directory, { withFileTypes: true })) {
		const path = join(directory, entry.name)

		if (entry.isDirectory()) sources(path, into)
		else if (/\.tsx?$/.test(entry.name)) into.push({ path, text: readFileSync(path, "utf8") })
	}

	return into
}

describe("strings the office editors share", () => {
	it("are none of DOCX editing's own in the spreadsheet viewer or the shared office components", () => {
		const docxOnly = new Set(
			sources(join(ROOT, "src/locales/en/previewDocx")).flatMap(({ text }) =>
				[...text.matchAll(/^\t(previewDocx[A-Za-z0-9]+?)(?:_one|_other)?:/gm)].map(match => match[1] ?? "")
			)
		)
		const used = [...sources(join(ROOT, "src/features/spreadsheet")), ...sources(join(ROOT, "src/features/office"))].flatMap(
			({ path, text }) =>
				[...text.matchAll(/["'`](previewDocx[A-Za-z0-9]+)["'`]/g)]
					.map(match => match[1] ?? "")
					.filter(key => docxOnly.has(key))
					.map(key => `${path.slice(ROOT.length)}: ${key}`)
		)

		expect(docxOnly.size).toBeGreaterThan(0)
		expect(used).toEqual([])
	})

	it("spell color as the rest of the app's English catalogue does", () => {
		const catalogs = [
			...sources(join(ROOT, "src/locales/en/previewOffice")),
			...sources(join(ROOT, "src/locales/en/previewSheet")),
			...sources(join(ROOT, "src/locales/en/previewDocx")),
			...sources(join(ROOT, "src/locales/en")).filter(({ path }) => path.endsWith("/preview.ts"))
		]
		const british = catalogs.flatMap(({ path, text }) =>
			[...text.matchAll(/^\t(\w+): "([^"]*\bcolours?\b[^"]*)"/gim)].map(match => `${path.slice(ROOT.length)}: ${match[1] ?? ""}`)
		)

		expect(british).toEqual([])
	})
})
