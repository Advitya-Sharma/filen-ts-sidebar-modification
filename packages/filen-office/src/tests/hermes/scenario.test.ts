import { readFileSync } from "node:fs"
import { describe, expect, it, vi } from "vitest"
import { decodeBase64, encodeBase64 } from "../../scripts/hermes/base64"
import { runScenario } from "../../scripts/hermes/scenario"

// The scenario src/scripts/hermes/run.ts runs under Hermes, here from source under V8: it must stay green
// without a Hermes CLI. CompressionStream is hidden, as in Hermes, so saves deflate through fflate.

function fixture(name: string): { name: string; bytes: Uint8Array } {
	return { name, bytes: new Uint8Array(readFileSync(new URL(`../fixtures/${name}`, import.meta.url))) }
}

describe("Hermes scenario under V8", () => {
	it("opens, lays out, edits, saves and reopens; unedited saves are the source", async () => {
		vi.stubGlobal("CompressionStream", undefined)

		const result = await runScenario([fixture("01-plain.docx"), fixture("04-comments-tracked.docx"), fixture("06-textbox-shape.docx")])

		expect(result.errors).toEqual([])
		expect(result.fixtures).toHaveLength(3)

		for (const done of result.fixtures) {
			expect(done.uneditedIdentical, done.name).toBe(true)
			expect(
				done.steps.every(step => step.endsWith(":true")),
				`${done.name}: ${done.steps.join()}`
			).toBe(true)
			expect(done.changedParts, done.name).toContain("word/document.xml")
			expect(done.reopenedPages, done.name).toBeGreaterThan(0)
			expect(done.reopenedMatches, done.name).toBe(1)
		}
	})

	it("runs the same twice: the result does not depend on timing", async () => {
		vi.stubGlobal("CompressionStream", undefined)

		const [first, second] = [await runScenario([fixture("05-fields-toc.docx")]), await runScenario([fixture("05-fields-toc.docx")])]

		expect(second.fixtures[0]?.structure).toBe(first.fixtures[0]?.structure)
		expect(second.fixtures[0]?.editedStructure).toBe(first.fixtures[0]?.editedStructure)
		expect(Array.from(second.fixtures[0]?.edited ?? [])).toEqual(Array.from(first.fixtures[0]?.edited ?? []))
	})

	it("round-trips base64", () => {
		const bytes = Uint8Array.from({ length: 1000 }, (_, k) => (k * 37) & 255)

		for (const length of [0, 1, 2, 3, 4, 999, 1000]) {
			const part = bytes.subarray(0, length)

			expect(encodeBase64(part)).toBe(Buffer.from(part).toString("base64"))
			expect(Array.from(decodeBase64(encodeBase64(part)))).toEqual(Array.from(part))
		}
	})
})
