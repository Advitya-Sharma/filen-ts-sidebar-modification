import { existsSync, mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { summarize } from "./fuzz/config"
import { ACCEPTED, ACCEPTED_COUNTS, acceptedCount, acceptedReason } from "./interop/accepted"
import { L3_GATED, runOffice, selectEntries, tally, type EntryResult } from "./interop/interop"
import { offices } from "./interop/soffice"

// Edited saves in LibreOffice (interop/interop.ts): opt-in with FILEN_OFFICE_SOFFICE=<soffice>[,<soffice>] and the
// saves the fuzzers kept (FILEN_OFFICE_FUZZ_SAVES, with manifest.json); skipped otherwise.
// FILEN_OFFICE_INTEROP_PER_FAMILY (50) saves per op family besides every scripted one, FILEN_OFFICE_INTEROP_JOBS
// (4) LibreOffice processes, FILEN_OFFICE_INTEROP_LIMIT caps the saves (a quick run), FILEN_OFFICE_INTEROP_WORK keeps the outputs and reports (else a temporary
// directory, removed).

const SAVES = process.env["FILEN_OFFICE_FUZZ_SAVES"] || undefined
const OFFICES = offices()
const ENABLED = SAVES !== undefined && existsSync(join(SAVES, "manifest.json")) && OFFICES.length > 0

// Failures of one entry, each "<check>: <detail>", less the accepted ones.
function failures(entry: EntryResult, office: string): string[] {
	const out: string[] = []

	for (const item of entry.l1) out.push(`L1 ${item}`)

	for (const view of ["final", "original"]) {
		const words = entry.l2[view]?.words

		if (words !== undefined && !words.equal) out.push(`L2 ${view}: ${JSON.stringify(words.residual)}`)
	}

	for (const [metric, delta] of Object.entries(entry.l3))
		if (
			L3_GATED.includes(metric) &&
			acceptedCount(metric, delta.theirs, delta.ours, entry.witness.before, entry.witness.after, office) === undefined
		)
			out.push(`L3 ${metric}: LibreOffice ${delta.theirs}, ours ${delta.ours}`)

	if (entry.g7.gate !== undefined) out.push(`G7 gate: ${entry.g7.gate}`)
	if (entry.g7.backGate !== undefined) out.push(`G7 back gate: ${entry.g7.backGate}`)
	if (entry.g7.text !== undefined && !entry.g7.text.equal) out.push(`G7 text: ${JSON.stringify(entry.g7.text.residual)}`)
	if (entry.g7.error !== undefined) out.push(`G7 threw: ${entry.g7.error}`)

	return out.filter(failure => acceptedReason(entry, failure, office) === undefined)
}

describe.skipIf(!ENABLED)("edited saves in LibreOffice", () => {
	const limit = process.env["FILEN_OFFICE_INTEROP_LIMIT"]
	const entries = ENABLED
		? selectEntries(
				SAVES as string,
				Number(process.env["FILEN_OFFICE_INTEROP_PER_FAMILY"] || "50"),
				limit === undefined || limit === "" ? undefined : Number(limit)
			)
		: []

	for (const office of OFFICES) {
		it(`converts and reads every edit the way we do (${office.name})`, { timeout: 14_400_000 }, async () => {
			const keep = process.env["FILEN_OFFICE_INTEROP_WORK"] || undefined
			const work = keep ?? mkdtempSync(join(tmpdir(), "filen-office-interop-"))

			try {
				const report = await runOffice(
					office,
					{
						saves: SAVES as string,
						entries,
						work,
						jobs: Number(process.env["FILEN_OFFICE_INTEROP_JOBS"] || "4"),
						timeoutMs: 60_000,
						batch: 12,
						layoutMaxBytes: 2_000_000
					},
					line => console.log(line)
				)
				const counts = tally(report)
				const failed = report.entries
					.map(entry => ({ entry: entry.path, failures: failures(entry, office.name) }))
					.filter(item => item.failures.length > 0)

				summarize(
					`LibreOffice ${office.version}: ${JSON.stringify(counts)}; ${failed.length} saves with unaccepted failures; ${ACCEPTED.length + ACCEPTED_COUNTS.length} accepted`
				)
				expect(counts.saves).toBeGreaterThan(0)
				expect(failed).toEqual([])
			} finally {
				if (keep === undefined) rmSync(work, { recursive: true, force: true })
			}
		})
	}
})
