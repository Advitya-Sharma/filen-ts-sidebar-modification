import { existsSync, readFileSync, readdirSync, writeFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { openOdtAsDocx } from "../core/odt/index"
import { bytesSource } from "@filen/office/core/opc"
import { CORPUS_DIR } from "./corpusDir"
import { layOut, pageLines } from "./layoutFixtures"
import { translationPackage } from "../core/odt/write/source"

// What the viewer shows of every corpus ODT, page by page: each line's text and place and a fingerprint of its glyph
// runs and graphics (not of source ids, which number the translation's items). FILEN_OFFICE_ODT_DIGESTS names a JSON
// file: written when absent, else compared with, so a translator change can be shown to leave every page as it was.
// Opt-in where the corpus exists.

const FILE = process.env["FILEN_OFFICE_ODT_DIGESTS"] || undefined
const DIRECTORY = join(CORPUS_DIR, "odt")

function hash(values: ArrayLike<number>): number {
	let out = 0

	for (let i = 0; i < values.length; i++) out = (Math.imul(out, 31) + Math.round((values[i] ?? 0) * 100)) | 0

	return out
}

describe.skipIf(FILE === undefined || !existsSync(DIRECTORY))("ODT viewer pages", () => {
	it("are what they were", { timeout: 3_600_000 }, async () => {
		const digests: Record<string, string[]> = {}

		for (const name of readdirSync(DIRECTORY)
			.filter(file => file.endsWith(".odt"))
			.sort()) {
			const opened = await openOdtAsDocx(bytesSource(new Uint8Array(readFileSync(join(DIRECTORY, name)))))

			if (!opened.ok) continue

			const layout = await layOut(await translationPackage(opened.translation))

			digests[name] = layout.pages.map((page, index) => {
				const lines = pageLines(layout, index)
					.map(line => `${line.text}@${Math.round(line.top * 100)},${Math.round(line.x * 100)},${Math.round(line.width * 100)}`)
					.join("\n")

				return `${lines}#${hash(page.text.fragmentFloats)}:${hash(page.graphicFloats)}:${hash(page.objectFloats)}`
			})
		}

		const file = FILE as string

		if (!existsSync(file)) {
			writeFileSync(file, JSON.stringify(digests))

			return
		}

		const before = JSON.parse(readFileSync(file, "utf8")) as Record<string, string[]>
		const changed = Object.keys({ ...before, ...digests }).filter(
			name => JSON.stringify(before[name]) !== JSON.stringify(digests[name])
		)

		expect(changed).toEqual([])
	})
})
