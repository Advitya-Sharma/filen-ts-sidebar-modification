import { appendFileSync, existsSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { CORPUS, corpusFiles, expectedRefusal, fuzzConfig, runJobs, summarize } from "./config"

// Random edit sequences through the paged editor over every corpus document: after each settled step the
// pages the editor shows (laid out incrementally) must equal a fresh layout of the file it saves.
// FILEN_OFFICE_FUZZ=<seeds per file and mode> turns it on; FILEN_OFFICE_FUZZ_SEED (first seed),
// FILEN_OFFICE_FUZZ_FILTER, FILEN_OFFICE_FUZZ_STEPS (20), FILEN_OFFICE_FUZZ_MODES (sync,race; also proof),
// FILEN_OFFICE_FUZZ_OPS, FILEN_OFFICE_FUZZ_VIEW, FILEN_OFFICE_FUZZ_JOBS and FILEN_OFFICE_FUZZ_OUT tune it.
// One run: FILEN_OFFICE_FUZZ=1 node --import tsx src/tests/fuzz/run.ts paged <file> <seed> <mode>

const config = fuzzConfig()
const JOB_TIMEOUT_MS = 1_800_000

describe.skipIf(config === undefined || !existsSync(CORPUS))("paged editing fuzz over the corpus", () => {
	it("shows what a fresh layout of the saved file shows", { timeout: 86_400_000 }, async () => {
		if (config === undefined) return

		const modes = config.modes ?? ["sync", "race"]
		const jobs: string[][] = []

		for (const file of corpusFiles(["docx"], config.filter))
			for (let seed = config.seed; seed < config.seed + config.runs; seed++)
				for (const mode of modes) jobs.push(["paged", file, String(seed), mode])

		expect(jobs.length, "no corpus file matches").toBeGreaterThan(0)

		const tally: Record<string, number> = {}
		const failures: string[] = []
		const started = performance.now()
		let slowest = { run: "", ms: 0 }

		await runJobs(jobs, config.jobs, JOB_TIMEOUT_MS, job => {
			const line = job.lines.at(-1)
			const run = job.args.slice(1).join(" ").replace(`${CORPUS}/`, "")

			if (config.out !== undefined) for (const output of job.lines) appendFileSync(config.out, `${output}\n`)

			if (line === undefined) {
				tally["noResult"] = (tally["noResult"] ?? 0) + 1
				failures.push(`${run}: ${job.timedOut ? "timed out" : `exit ${job.code}`} ${job.stderr.slice(-600)}`)
				return
			}

			const result = JSON.parse(line) as Record<string, unknown>
			const kind =
				result["ok"] === true
					? "ok"
					: typeof result["refused"] === "string"
						? result["refused"]
						: (["diverge", "crash", "saveFail"].find(key => key in result) ?? "unknown")

			tally[kind] = (tally[kind] ?? 0) + 1

			if (typeof result["ms"] === "number" && result["ms"] > slowest.ms) slowest = { run, ms: result["ms"] }
			if (kind !== "ok" && !expectedRefusal(kind)) failures.push(line.slice(0, 3000))
			// An editable document no op managed to change was not fuzzed at all.
			if (kind === "ok" && result["applied"] === 0) failures.push(`${run}: no edit applied`)
		})

		summarize(
			`paged fuzz: ${jobs.length} runs (seeds ${config.seed}-${config.seed + config.runs - 1}, ${modes.join("+")}) ${JSON.stringify(tally)}; ` +
				`${((performance.now() - started) / 1000).toFixed(0)} s on ${config.jobs} processes; slowest ${slowest.run} ${slowest.ms} ms`
		)
		expect(failures).toEqual([])
		expect(tally["ok"] ?? 0, "no run got to edit").toBeGreaterThan(0)
	})
})
