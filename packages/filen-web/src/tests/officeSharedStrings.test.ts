import { readFileSync, readdirSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import * as off from "@/locales/en/previewOffice/off"

// Strings the spreadsheet viewer on the office engine shares with DOCX editing live in a catalogue both flags ship
// (locales/en/previewOffice/*): a build with only one of them shows no raw key.

const ROOT = process.cwd()
const LOCALES = join(ROOT, "src/locales/en/previewOffice")

function sources(directory: string, into: { path: string; text: string }[] = []): { path: string; text: string }[] {
	for (const entry of readdirSync(directory, { withFileTypes: true })) {
		const path = join(directory, entry.name)

		if (entry.isDirectory()) sources(path, into)
		else if (/\.tsx?$/.test(entry.name)) into.push({ path, text: readFileSync(path, "utf8") })
	}

	return into
}

function aliases(): [RegExp, string][] {
	const config = readFileSync(join(ROOT, "vite.config.ts"), "utf8")
	const body = config.slice(config.indexOf("function officeSharedAlias("), config.indexOf("const contentSecurityPolicy"))

	return [...body.matchAll(/find: \/(.+?)\/,\s*replacement: "([^"]+)"/g)].map(match => [new RegExp(match[1] ?? ""), match[2] ?? ""])
}

describe("strings the office editors share", () => {
	it("are none of DOCX editing's own in the spreadsheet viewer or the shared office components", () => {
		const docxOnly = new Set(
			sources(join(ROOT, "src/locales/en/previewDocx"))
				.filter(({ path }) => !path.endsWith("off.ts"))
				.flatMap(({ text }) => [...text.matchAll(/^\t(previewDocx[A-Za-z0-9]+?)(?:_one|_other)?:/gm)].map(match => match[1] ?? ""))
		)
		const used = [...sources(join(ROOT, "src/features/spreadsheet/next")), ...sources(join(ROOT, "src/features/office"))].flatMap(
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

	it("are registered in the preview catalog, and left out of builds with neither flag", async () => {
		expect(readFileSync(join(ROOT, "src/locales/en/preview.ts"), "utf8")).toContain("...previewOfficeShared")

		const found = aliases()
		const files = readdirSync(LOCALES).filter(file => file.endsWith(".ts") && file !== "off.ts")

		expect(files.length).toBeGreaterThan(0)

		for (const file of files) {
			const specifier = `@/locales/en/previewOffice/${file.replace(/\.ts$/, "")}`

			expect(found.find(([pattern]) => pattern.test(specifier))?.[1], specifier).toBe("/src/locales/en/previewOffice/off.ts")

			const catalog = (await import(/* @vite-ignore */ join(LOCALES, file))) as Record<string, object>

			for (const name of Object.keys(catalog)) expect((off as Record<string, object>)[name], `${file} ${name}`).toEqual({})
		}
	})
})
