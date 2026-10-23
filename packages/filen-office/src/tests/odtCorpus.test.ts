import { existsSync, readdirSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { bytesSource } from "@filen/office/core/opc"
import { translateOdt } from "../core/odt/index"
import { CORPUS_DIR } from "./corpusDir"
import { layOutOdt } from "./odtFixtures"

// The licensed ODT corpus against LibreOffice 7.5's own rendering (fixtures/odt/libreoffice-pages.json: its
// PDF export's page count per file, a file it cannot load left out).
const DIRECTORY = join(CORPUS_DIR, "odt")
const REFERENCE = JSON.parse(readFileSync(new URL("./fixtures/odt/libreoffice-pages.json", import.meta.url), "utf8")) as {
	readonly pages: Readonly<Record<string, number>>
}

describe.skipIf(!existsSync(DIRECTORY))("the ODT corpus", () => {
	const files = existsSync(DIRECTORY)
		? readdirSync(DIRECTORY)
				.filter(name => name.endsWith(".odt"))
				.sort()
		: []

	it("opens at least every file LibreOffice opens, and refuses the rest with a reason", async () => {
		let opened = 0

		for (const name of files) {
			const result = await translateOdt(bytesSource(new Uint8Array(readFileSync(join(DIRECTORY, name)))))

			if (result.ok) opened++
			else expect(result.reason).toMatch(/^[a-zA-Z]+$/)
		}

		expect(opened).toBeGreaterThanOrEqual(Object.keys(REFERENCE.pages).length)
	}, 120_000)

	it("lays out within one page of LibreOffice on at least 85 % of the files", async () => {
		let within = 0
		let compared = 0

		for (const name of files) {
			const expected = REFERENCE.pages[name.replace(/\.odt$/, "")]

			if (expected === undefined) continue

			const { pages } = await layOutOdt(new Uint8Array(readFileSync(join(DIRECTORY, name))))

			compared++

			if (Math.abs(pages - expected) <= 1) within++
		}

		expect(compared).toBeGreaterThan(0)
		expect(within / compared).toBeGreaterThanOrEqual(0.85)
	}, 300_000)
})
