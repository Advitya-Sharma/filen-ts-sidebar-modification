import { appendFileSync, existsSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { CORPUS, corpusFiles, expectedRefusal, fuzzConfig, runJobs, summarize } from "./config"

// Raw edit sequences over every editable corpus document (body and, at random, its other stories), saved
// through Gates 1 and 2 and read back by the independent reader: a save Gate 2 lets through must open,
// keep untouched entries byte for byte, add nothing Word refuses, and hold exactly the editor's paragraphs
// among the source's. Refusals by the gates are intended and only counted.
// FILEN_OFFICE_FUZZ=<sequences per file> turns it on; the first sequence of a file saves unedited.
// FILEN_OFFICE_FUZZ_SEED, FILEN_OFFICE_FUZZ_FILTER, FILEN_OFFICE_FUZZ_STEPS (most ops per session, 16),
// FILEN_OFFICE_FUZZ_MODES (stories; also undo, blockmove, undoall, xsd, selftest), FILEN_OFFICE_FUZZ_OPS,
// FILEN_OFFICE_FUZZ_JOBS and FILEN_OFFICE_FUZZ_OUT tune it.
// One file: FILEN_OFFICE_FUZZ=8 node --import tsx src/tests/fuzz/run.ts damage <file>

const config = fuzzConfig()
const JOB_TIMEOUT_MS = 1_800_000

describe.skipIf(config === undefined || !existsSync(CORPUS))("writer damage fuzz over the corpus", () => {
	it("lets nothing through Gate 2 that the independent reader rejects", { timeout: 86_400_000 }, async () => {
		if (config === undefined) return

		const jobs = corpusFiles(["docx", "docm"], config.filter).map(file => ["damage", file])

		expect(jobs.length, "no corpus file matches").toBeGreaterThan(0)
		const stats: Record<string, number> = {}
		const findings: Record<string, number> = {}
		const refused: Record<string, number> = {}
		const failures: string[] = []
		const started = performance.now()
		let files = 0
		let slowest = { file: "", ms: 0 }

		await runJobs(jobs, config.jobs, JOB_TIMEOUT_MS, job => {
			let summary: Record<string, unknown> | undefined

			for (const line of job.lines) {
				if (config.out !== undefined) appendFileSync(config.out, `${line}\n`)

				const value = JSON.parse(line) as Record<string, unknown>
				const kind = value["finding"]

				if (typeof kind !== "string") {
					summary = value
					continue
				}

				findings[kind] = (findings[kind] ?? 0) + 1

				if (!kind.startsWith("refused:")) failures.push(line.slice(0, 3000))
			}

			if (summary === undefined) {
				failures.push(`${job.args[1]}: ${job.timedOut ? "timed out" : `exit ${job.code}`} ${job.stderr.slice(-600)}`)
				return
			}

			if (typeof summary["refused"] === "string") {
				refused[summary["refused"]] = (refused[summary["refused"]] ?? 0) + 1

				if (!expectedRefusal(summary["refused"])) failures.push(`${String(summary["file"])}: ${summary["refused"]}`)

				return
			}

			files++

			const fileStats = (summary["stats"] ?? {}) as Record<string, number>

			// An editable document no op managed to change was not fuzzed at all.
			if ((fileStats["seq"] ?? 0) > 1 && (fileStats["applied"] ?? 0) === 0)
				failures.push(`${String(summary["file"])}: no edit applied`)

			for (const [key, value] of Object.entries(fileStats)) stats[key] = (stats[key] ?? 0) + value

			if (typeof summary["ms"] === "number" && summary["ms"] > slowest.ms)
				slowest = { file: String(summary["file"]), ms: summary["ms"] }
		})

		summarize(
			`damage fuzz: ${files} files, ${config.runs} sequences each (seed ${config.seed}, ${(config.modes ?? ["stories"]).join("+")}) ` +
				`${JSON.stringify(stats)}; findings ${JSON.stringify(findings)}; not fuzzed ${JSON.stringify(refused)}; ${((performance.now() - started) / 1000).toFixed(0)} s on ` +
				`${config.jobs} processes; slowest ${slowest.file} ${slowest.ms} ms`
		)
		expect(failures).toEqual([])
		expect(stats["verified"] ?? 0, "no sequence was verified").toBeGreaterThan(0)
	})
})
