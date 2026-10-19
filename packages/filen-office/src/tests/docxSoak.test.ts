import { existsSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { DocxEditorSession } from "@filen/office/core/docx/edit"
import { soak, soakFailures, type SoakResult, type SoakSave } from "../bench/soak"
import { CORPUS_DIR } from "./corpusDir"
import { openBytes } from "./docxEditFixtures"

// Long editing sessions (opt-in: FILEN_OFFICE_SOAK=<minutes>, the plan run again until that much time went by,
// once at least): memory after a full collection (compiled code left out), in the heap and outside it, grows with
// the content only and stays within a bound of the open's, the undo history holds no more than a copy of the
// content, keys do not get slower, a save costs what was edited since the open, no save fails, and the last save
// reopens with the session's text (soakFailures).

const MINUTES = Number(process.env["FILEN_OFFICE_SOAK"] ?? "")
const ENABLED = Number.isFinite(MINUTES) && MINUTES > 0
const FIXTURE = join(import.meta.dirname, "fixtures", "08-large-300p.docx")
const TARIFF = join(CORPUS_DIR, "docx", "govuk__Tariff_of_the_United_Kingdom_v1.34.docx")

async function run(path: string, label: string, growthMb: number): Promise<void> {
	const result = await soak(new Uint8Array(readFileSync(path)), {
		minutes: MINUTES,
		onSave: save =>
			console.log(
				`${label} save after ${String(save.operations)} operations: heap ${save.heapMb.toFixed(1)} MB (raw ${save.rawHeapMb.toFixed(1)}, ` +
					`compiled code ${save.codeMb.toFixed(1)}, content ${save.contentMb.toFixed(1)}), array buffers ` +
					`${save.arrayBuffersMb.toFixed(1)} MB, save ${save.saveMs.toFixed(0)} ms (Gate 2 ${save.gate2Ms.toFixed(0)} ms, ` +
					`longest task ${save.longestTaskMs.toFixed(0)} ms), keys p95 ${save.keysP95.toFixed(2)} ms`
			)
	})
	const failures = soakFailures(result, { growthMb })

	console.log(
		`${label}: ${String(result.passes)} passes, ${String(result.keys.length)} keys (${String(result.refused)} refused), ` +
			`${String(result.saves.length)} saves, open heap ${result.openHeapMb.toFixed(1)} MB, after the last save ` +
			`${(result.saves[result.saves.length - 1]?.heapMb ?? 0).toFixed(1)} MB with the undo history and ` +
			`${result.clearedHeapMb.toFixed(1)} MB without (it held ${result.historyMb.toFixed(1)} of at most ${result.historyLimitMb.toFixed(1)} MB); ` +
			`${failures.length === 0 ? "within the limits" : failures.join("; ")}`
	)

	if (result.bytes === undefined) throw new Error("nothing saved")

	const reopened = await openBytes(result.bytes)
	const session = new DocxEditorSession(reopened.context)

	expect(session.state.doc.textContent).toBe(result.text)
	session.dispose()
	expect(failures).toEqual([])
}

describe.skipIf(!ENABLED)("docx editing soak", () => {
	it("keeps a long session on the 300-page fixture flat", { timeout: 0 }, async () => {
		await run(FIXTURE, "300p", 16)
	})

	it.skipIf(!existsSync(TARIFF))("keeps a long session on Tariff flat", { timeout: 0 }, async () => {
		await run(TARIFF, "Tariff", 64)
	})
})

describe("the soak's limits", () => {
	const result = (save: (k: number) => Partial<SoakSave>): SoakResult => ({
		openHeapMb: 100,
		openContentMb: 20,
		clearedHeapMb: 110,
		historyMb: 0,
		historyLimitMb: 10,
		saves: Array.from({ length: 20 }, (_, k): SoakSave => ({
			operations: k * 250,
			heapMb: 110,
			rawHeapMb: 120,
			codeMb: 10,
			contentMb: 20,
			externalMb: 302,
			arrayBuffersMb: 300,
			saveMs: 100,
			saveCpuMs: 90,
			editedMb: 1,
			gate2Ms: 10,
			longestTaskMs: 10,
			longestTaskCpuMs: 10,
			keysP95: 1,
			...save(k)
		})),
		keys: Array.from({ length: 100 }, () => 1),
		refused: 0,
		failedSaves: [],
		passes: 1,
		bytes: undefined,
		text: ""
	})
	const limits = { growthMb: 64 }

	it("pass a session whose memory and saves follow its content", () => {
		expect(
			soakFailures(
				result(() => ({})),
				limits
			)
		).toEqual([])
		// Content growing 1 MB per save: heap, array buffers and save CPU with it.
		expect(
			soakFailures(
				result(k => ({
					contentMb: 20 + k,
					heapMb: 110 + k * 1.2,
					arrayBuffersMb: 300 + k * 6,
					editedMb: 1 + k,
					saveCpuMs: 90 * (1 + k)
				})),
				limits
			)
		).toEqual([])
	})

	it("fail on array buffers kept from each save, the heap flat", () => {
		expect(
			soakFailures(
				result(k => ({ arrayBuffersMb: 300 + k * 8 })),
				limits
			)
		).toEqual(["array buffers grow 8.00 MB per save over the last 15 saves (content 0.00)"])
	})

	it("fail on the heap growing apart from the content, and on saves costing more for the same edits", () => {
		expect(
			soakFailures(
				result(k => ({ heapMb: 110 + k * 2 })),
				limits
			)
		).toEqual(["heap less content grows 2.00 MB per save over the last 15 saves"])
		expect(
			soakFailures(
				result(k => ({ saveCpuMs: 90 + k * 10 })),
				limits
			)
		).toEqual(["last save 280 ms CPU per edited MB over 1.25 x the second's 100"])
		expect(
			soakFailures(
				result(k => ({ longestTaskCpuMs: k === 7 ? 80 : 10 })),
				limits
			)
		).toEqual(["a save task took 80 ms CPU"])
	})
})
