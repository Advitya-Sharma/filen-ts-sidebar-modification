import { existsSync, readFileSync, statSync, writeFileSync } from "node:fs"
import { availableParallelism } from "node:os"
import { describe, expect, it } from "vitest"
import { CORPUS, corpusFiles, runJobs, summarize } from "./config"

// The viewer's layout of every unedited corpus document, digested, against a baseline: a change meant to
// leave the viewer alone must leave every digest alone. FILEN_OFFICE_DIGEST=<file.json> turns it on: the
// first run (before the change) writes the baseline there, later runs compare with it.
// FILEN_OFFICE_FUZZ_JOBS sets the processes (default: one per core but one).

const BASELINE = process.env["FILEN_OFFICE_DIGEST"] || undefined
const MAX_BYTES = 8_000_000
const FILES_PER_JOB = 20

describe.skipIf(BASELINE === undefined || !existsSync(CORPUS))("viewer layout digests of the corpus", () => {
	it("lays out every unedited document as the baseline did", { timeout: 3_600_000 }, async () => {
		if (BASELINE === undefined) return

		const files = corpusFiles(["docx"]).filter(file => statSync(file).size < MAX_BYTES)

		expect(files.length, "no corpus files").toBeGreaterThan(0)
		const jobs: string[][] = []
		const current: Record<string, string> = {}
		const failures: string[] = []

		for (let i = 0; i < files.length; i += FILES_PER_JOB) jobs.push(["viewer", ...files.slice(i, i + FILES_PER_JOB)])

		await runJobs(jobs, Number(process.env["FILEN_OFFICE_FUZZ_JOBS"]) || Math.max(1, availableParallelism() - 1), 1_800_000, job => {
			for (const line of job.lines) {
				const value = JSON.parse(line) as { file: string; digest: string }

				current[value.file] = value.digest
			}

			if (job.lines.length !== job.args.length - 1)
				failures.push(`${job.args[1]}...: ${job.timedOut ? "timed out" : `exit ${job.code}`}`)
		})

		expect(failures).toEqual([])

		const sorted = Object.fromEntries(Object.entries(current).sort(([a], [b]) => (a < b ? -1 : 1)))

		if (!existsSync(BASELINE)) {
			writeFileSync(BASELINE, `${JSON.stringify(sorted, null, 1)}\n`)
			summarize(`viewer digests: baseline of ${files.length} files written to ${BASELINE}`)
			return
		}

		const baseline = JSON.parse(readFileSync(BASELINE, "utf8")) as Record<string, string>
		const changes: string[] = []

		for (const [file, digest] of Object.entries(sorted))
			if (baseline[file] !== digest) changes.push(`${file}: ${baseline[file] ?? "new"} -> ${digest}`)
		for (const file of Object.keys(baseline)) if (sorted[file] === undefined) changes.push(`${file}: gone`)

		summarize(`viewer digests: ${changes.length}/${files.length} differ from ${BASELINE}`)
		expect(changes).toEqual([])
	})
})
