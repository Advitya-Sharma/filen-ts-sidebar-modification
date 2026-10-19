import { existsSync } from "node:fs"
import { availableParallelism } from "node:os"
import { describe, expect, it } from "vitest"
import { CORPUS, WORD_EXTENSIONS, corpusFiles, expectedRefusal, runJobs, summarize } from "./config"
import { savesConfig, writeManifest } from "./saves"

// The scripted edit suite (every op's deterministic edit) over every editable corpus document, saved under
// FILEN_OFFICE_FUZZ_SAVES for the file oracles; skipped without it. FILEN_OFFICE_FUZZ_FILTER and
// FILEN_OFFICE_FUZZ_JOBS tune it.

const saves = savesConfig()
const FILES_PER_JOB = 10

describe.skipIf(saves === undefined || !existsSync(CORPUS))("scripted edit suite saves", () => {
	it("saves the scripted edits of every editable document", { timeout: 3_600_000 }, async () => {
		if (saves === undefined) return

		const files = corpusFiles(WORD_EXTENSIONS, process.env["FILEN_OFFICE_FUZZ_FILTER"] || undefined)
		const jobs: string[][] = []

		for (let i = 0; i < files.length; i += FILES_PER_JOB) jobs.push(["script", ...files.slice(i, i + FILES_PER_JOB)])

		expect(jobs.length, "no corpus file matches").toBeGreaterThan(0)

		const tally: Record<string, number> = {}
		const failures: string[] = []
		const bump = (key: string): void => {
			tally[key] = (tally[key] ?? 0) + 1
		}

		await runJobs(jobs, Number(process.env["FILEN_OFFICE_FUZZ_JOBS"]) || Math.max(1, availableParallelism() - 1), 1_800_000, job => {
			for (const line of job.lines) {
				const value = JSON.parse(line) as Record<string, unknown>

				if (value["ok"] === true) bump(value["kept"] === true ? "kept" : "unchanged")
				else if (typeof value["refused"] === "string") {
					bump(value["refused"])

					if (!expectedRefusal(value["refused"])) failures.push(line)
				} else {
					bump(typeof value["saveRefused"] === "string" ? `saveRefused:${value["saveRefused"]}` : "crash")
					failures.push(line.slice(0, 2000))
				}
			}

			if (job.lines.length !== job.args.length - 1)
				failures.push(`${job.args[1]}...: ${job.timedOut ? "timed out" : `exit ${job.code}`}`)
		})

		summarize(`scripted saves: ${JSON.stringify(tally)}; ${writeManifest(saves.dir)} saves in ${saves.dir}`)
		expect(failures).toEqual([])
	})
})
