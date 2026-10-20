import { appendFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { summarize } from "./fuzz/config"
import { fixtureBytes, FIXTURES } from "./fuzz/sheet/fixtures"
import { DEFAULT_SEQUENCE, isFinding, runSequence, type FuzzMode, type SequenceResult } from "./fuzz/sheet/harness"
import { runSheetJobs, sheetFuzzConfig, sheetFuzzTargets } from "./fuzz/sheet/run"

// Random command sequences through xlsx sessions, awaited one by one ("sync") and sent in bursts that interleave
// with tile requests, copies and saves ("race"). At every checkpoint the save passes its gate, reads back without
// the engine as the session's model holds it, and a fresh session on it shows what the editing session shows; once
// everything is undone the save gives the source back. Refusals are typed and tallied, never findings.
// The quick run (the fixtures, in this process) is always on. FILEN_OFFICE_SHEET_FUZZ=<sequences> adds the long run
// over the fixtures and the corpus, each workbook's seeds in a process of their own (fuzz/sheet/README.md).

const QUICK_SEEDS = 6
const MODES: readonly FuzzMode[] = ["sync", "race"]
const JOB_TIMEOUT_MS = 1_800_000

function tallyRefusals(into: Record<string, number>, result: SequenceResult): void {
	for (const [reason, times] of Object.entries(result.refusals)) into[reason] = (into[reason] ?? 0) + times
}

describe("sheet fuzz over the fixtures", () => {
	for (const name of Object.keys(FIXTURES))
		for (const mode of MODES)
			it(`${name}, ${mode}`, { timeout: 120_000 }, async () => {
				const bytes = fixtureBytes(name)
				const findings: SequenceResult[] = []
				const refusals: Record<string, number> = {}
				let applied = 0

				for (let seed = 1; seed <= QUICK_SEEDS; seed++) {
					// Cells at the grid's far edges and range-valued formulas run in the long run only, in processes of
					// their own.
					const result = await runSequence(bytes, "xlsx", name, seed, mode, { ...DEFAULT_SEQUENCE, edges: false, ranges: false })

					applied += result.applied
					tallyRefusals(refusals, result)

					if (isFinding(result) || result.skipped !== undefined) findings.push(result)
				}

				expect(findings.map(result => JSON.stringify(result).slice(0, 3000))).toEqual([])
				expect(applied).toBeGreaterThan(QUICK_SEEDS * 4)
			})
})

const config = sheetFuzzConfig()

describe.skipIf(config === undefined)("sheet fuzz, long run", () => {
	it("keeps every save equal to its session", { timeout: 86_400_000 }, async () => {
		if (config === undefined) return

		const targets = sheetFuzzTargets(config.filter, config.maxBytes)

		expect(targets.length, "no workbook matches").toBeGreaterThan(0)

		const perTarget = Math.max(1, Math.ceil(config.sequences / (targets.length * config.modes.length)))
		const jobs: string[][] = []

		for (const target of targets) for (const mode of config.modes) jobs.push([target, String(config.seed), String(perTarget), mode])

		const tally: Record<string, number> = {}
		const refusals: Record<string, number> = {}
		const skipped: Record<string, number> = {}
		const failures: string[] = []
		const started = performance.now()
		let sequences = 0
		let slowest = { run: "", ms: 0 }

		await runSheetJobs(jobs, config.jobs, JOB_TIMEOUT_MS, job => {
			const run = job.args.join(" ")

			if (config.out !== undefined) for (const line of job.lines) appendFileSync(config.out, `${line}\n`)

			// A job that died before its last sequence (a crash outside a sequence, an out-of-memory, a hang).
			if (job.code !== 0) {
				tally["jobDied"] = (tally["jobDied"] ?? 0) + 1
				failures.push(
					`${run}: ${job.timedOut ? "timed out" : `exit ${job.code}`} after ${job.lines.length} sequences ${job.stderr.slice(-800)}`
				)
			}

			for (const line of job.lines) {
				const result = JSON.parse(line) as SequenceResult

				sequences++
				tallyRefusals(refusals, result)

				if (result.ms > slowest.ms) slowest = { run: `${result.name} ${result.seed} ${result.mode}`, ms: result.ms }

				if (result.skipped !== undefined) {
					skipped[result.skipped] = (skipped[result.skipped] ?? 0) + 1
					continue
				}

				const kind =
					(["diverge", "saveFail", "untypedRefusal", "crash", "hang", "undoAll", "reopen"] as const).find(
						key => result[key] !== undefined
					) ?? "ok"

				tally[kind] = (tally[kind] ?? 0) + 1

				if (kind !== "ok") failures.push(line.slice(0, 3000))
			}
		})

		summarize(
			`sheet fuzz: ${sequences} sequences over ${targets.length} workbooks (${config.modes.join("+")}, ${config.steps} steps, ` +
				`edges ${config.edges}, ranges ${config.ranges}) ${JSON.stringify(tally)}; skipped ${JSON.stringify(skipped)}; ` +
				`refused ${JSON.stringify(refusals)}; ${((performance.now() - started) / 1000).toFixed(0)} s on ${config.jobs} processes; ` +
				`slowest ${slowest.run} ${slowest.ms} ms`
		)
		expect(failures).toEqual([])
		expect(tally["ok"] ?? 0).toBeGreaterThan(0)
	})
})
