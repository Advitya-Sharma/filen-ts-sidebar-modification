import { existsSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { openOdtSource } from "../core/odt/write/source"
import { proveOdt, writeOdt } from "../core/odt/write/writer"
import { CORPUS_DIR } from "./corpusDir"
import { decodeUtf8, encodeUtf8 } from "@filen/office/core/xml"
import { budget, cpuMs } from "./timing"

// The largest corpus ODT (14 MB, pictures mostly) read, proved and saved where the web runs it (the docx worker):
// CPU time per step, peak memory growth, and the longest stretch the worker goes without a turn when the steps run
// in 12 ms slices (a close must be read within it). Opt-in where the corpus exists.

const FILE = join(CORPUS_DIR, "odt", "govuk__8_Energy_20_5_2026v2.odt")

function turns(): {
	turns: { signal: { aborted: boolean }; slicing: { sliceMs: number; now: () => number; pause: () => Promise<void> } }
	longest: () => number
} {
	let last = cpuMs()
	let longest = 0

	return {
		turns: {
			signal: { aborted: false },
			slicing: {
				sliceMs: 12,
				now: () => performance.now(),
				pause: () => {
					longest = Math.max(longest, cpuMs() - last)

					return new Promise<void>(resolve =>
						setImmediate(() => {
							last = cpuMs()
							resolve()
						})
					)
				}
			}
		},
		longest: () => Math.max(longest, cpuMs() - last)
	}
}

describe.skipIf(!existsSync(FILE))("ODT writer on a 14 MB document", () => {
	it("reads, proves and saves within the docx save budgets", { timeout: 600_000 }, async () => {
		const bytes = new Uint8Array(readFileSync(FILE))
		const rss = process.resourceUsage().maxRSS
		let started = cpuMs()
		const opening = turns()
		const opened = await openOdtSource(bytes, opening.turns)
		const openMs = cpuMs() - started

		if (!opened.ok) throw new Error(opened.message)

		const openLongest = opening.longest()

		started = cpuMs()

		const proving = turns()
		const proof = await proveOdt(opened.source, proving.turns)
		const proveMs = cpuMs() - started
		const proveLongest = proving.longest()
		// A typed edit as the docx writer writes it: the body part, its first text grown.
		const body = decodeUtf8(opened.source.translation.parts.get("word/document.xml") as Uint8Array)
		const at = body.indexOf(">", body.indexOf("<w:t ")) + 1
		const edits = { parts: new Map([["word/document.xml", encodeUtf8(`${body.slice(0, at)}Filen ${body.slice(at)}`)]]) }

		started = cpuMs()

		const saving = turns()
		const written = await writeOdt(opened.source, edits, { turns: saving.turns })
		const saveMs = cpuMs() - started
		const saveLongest = saving.longest()
		const grown = (process.resourceUsage().maxRSS - rss) / 1024

		console.log(
			`open ${openMs.toFixed(0)} ms (longest stretch ${openLongest.toFixed(0)} ms), prove ${proveMs.toFixed(0)} ms (${proveLongest.toFixed(0)} ms, ${proof.locked.length} locked), save ${saveMs.toFixed(0)} ms (${saveLongest.toFixed(0)} ms, ${written.written} written, ${written.bytes.length} bytes), peak RSS +${grown.toFixed(0)} MB`
		)

		expect(written.bytes.length).toBeGreaterThan(bytes.length / 2)
		// The docx editor's save budget for a document this size; no task of the save or the proof past 50 ms.
		expect(saveMs).toBeLessThan(budget(5000))
		expect(saveLongest).toBeLessThan(budget(50))
		expect(proveLongest).toBeLessThan(budget(50))
	})
})
