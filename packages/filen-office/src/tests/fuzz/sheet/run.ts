import { spawn } from "node:child_process"
import { existsSync, readFileSync, readdirSync, statSync } from "node:fs"
import { availableParallelism } from "node:os"
import { basename, join } from "node:path"
import type { SheetFileKind } from "../../../core/sheet/protocol"
import { CORPUS } from "../config"
import { FIXTURES, fixtureBytes } from "./fixtures"
import { DEFAULT_SEQUENCE, runSequence, type FuzzMode, type SequenceOptions } from "./harness"

// The sheet fuzz's long runs: each job is one workbook, a range of seeds and a mode, in a Node process of its own
// (a crash or a hang costs that job alone). Prints one JSON line per sequence.
//   node --import tsx src/tests/fuzz/sheet/run.ts <corpus file | fixture:<name>> <first seed> <count> <sync|race>
// Reads FILEN_OFFICE_SHEET_FUZZ_* (sheetFuzzConfig) for steps, families, ops, edges and ranges.

export interface SheetFuzzConfig {
	// Sequences in all, spread over the workbooks and modes.
	readonly sequences: number
	readonly seed: number
	readonly filter: string | undefined
	readonly steps: number
	readonly modes: readonly FuzzMode[]
	readonly families: readonly string[] | undefined
	readonly ops: readonly string[] | undefined
	readonly edges: boolean
	readonly ranges: boolean
	readonly jobs: number
	// Corpus workbooks larger than this are left out (each checkpoint saves and reads the whole file).
	readonly maxBytes: number
	readonly out: string | undefined
}

function list(value: string | undefined): readonly string[] | undefined {
	return value === undefined || value === "" ? undefined : value.split(",").map(item => item.trim())
}

function whole(value: string | undefined, name: string, fallback: number): number {
	if (value === undefined || value === "") return fallback

	const parsed = Number(value)

	if (!Number.isInteger(parsed) || parsed < 0) throw new Error(`${name} must be a whole number: ${value}`)

	return parsed
}

// FILEN_OFFICE_SHEET_FUZZ=<sequences> turns the long run on.
export function sheetFuzzConfig(env: NodeJS.ProcessEnv = process.env): SheetFuzzConfig | undefined {
	const sequences = whole(env["FILEN_OFFICE_SHEET_FUZZ"], "FILEN_OFFICE_SHEET_FUZZ", 0)

	if (sequences === 0) return undefined

	const modes = (list(env["FILEN_OFFICE_SHEET_FUZZ_MODES"]) ?? ["sync", "race"]).map(mode => {
		if (mode !== "sync" && mode !== "race") throw new Error(`FILEN_OFFICE_SHEET_FUZZ_MODES: sync or race, not ${mode}`)

		return mode
	})

	return {
		sequences,
		seed: whole(env["FILEN_OFFICE_SHEET_FUZZ_SEED"], "FILEN_OFFICE_SHEET_FUZZ_SEED", 1),
		filter: env["FILEN_OFFICE_SHEET_FUZZ_FILTER"] || undefined,
		steps: whole(env["FILEN_OFFICE_SHEET_FUZZ_STEPS"], "FILEN_OFFICE_SHEET_FUZZ_STEPS", DEFAULT_SEQUENCE.steps),
		modes,
		families: list(env["FILEN_OFFICE_SHEET_FUZZ_FAMILIES"]),
		ops: list(env["FILEN_OFFICE_SHEET_FUZZ_OPS"]),
		edges: env["FILEN_OFFICE_SHEET_FUZZ_EDGES"] !== "0",
		ranges: env["FILEN_OFFICE_SHEET_FUZZ_RANGES"] !== "0",
		jobs: whole(
			env["FILEN_OFFICE_SHEET_FUZZ_JOBS"],
			"FILEN_OFFICE_SHEET_FUZZ_JOBS",
			Math.min(4, Math.max(1, availableParallelism() - 1))
		),
		maxBytes: whole(env["FILEN_OFFICE_SHEET_FUZZ_MAX_BYTES"], "FILEN_OFFICE_SHEET_FUZZ_MAX_BYTES", 5_000_000),
		out: env["FILEN_OFFICE_SHEET_FUZZ_OUT"] || undefined
	}
}

export function sequenceOptions(config: Pick<SheetFuzzConfig, "steps" | "families" | "ops" | "edges" | "ranges">): SequenceOptions {
	return {
		...DEFAULT_SEQUENCE,
		steps: config.steps,
		edges: config.edges,
		ranges: config.ranges,
		...(config.families !== undefined && { families: config.families }),
		...(config.ops !== undefined && { ops: config.ops })
	}
}

const KINDS: readonly SheetFileKind[] = ["xlsx", "xlsm", "xltx", "xltm"]

// The fixtures and the corpus spreadsheets (at most `maxBytes`), matching `filter`.
export function sheetFuzzTargets(filter: string | undefined, maxBytes: number): string[] {
	const targets = Object.keys(FIXTURES).map(name => `fixture:${name}`)

	for (const kind of KINDS) {
		const directory = join(CORPUS, kind)

		if (!existsSync(directory)) continue

		for (const name of readdirSync(directory).sort()) {
			const path = join(directory, name)

			if (name.endsWith(`.${kind}`) && statSync(path).size <= maxBytes) targets.push(path)
		}
	}

	return filter === undefined ? targets : targets.filter(target => target.includes(filter))
}

function target(name: string): { bytes: Uint8Array; kind: SheetFileKind; label: string } {
	if (name.startsWith("fixture:")) {
		const fixture = name.slice("fixture:".length)

		return { bytes: fixtureBytes(fixture), kind: "xlsx", label: name }
	}

	const extension = name.slice(name.lastIndexOf(".") + 1) as SheetFileKind

	return {
		bytes: new Uint8Array(readFileSync(name)),
		kind: KINDS.includes(extension) ? extension : "xlsx",
		label: name.startsWith(CORPUS) ? name.slice(CORPUS.length + 1) : basename(name)
	}
}

export interface SheetJobResult {
	readonly args: readonly string[]
	readonly lines: readonly string[]
	readonly code: number | null
	readonly timedOut: boolean
	readonly stderr: string
}

const RUNNER = join(import.meta.dirname, "run.ts")
const PACKAGE = join(import.meta.dirname, "..", "..", "..", "..")
const running = new Set<ReturnType<typeof spawn>>()

process.once("exit", () => {
	for (const child of running) child.kill("SIGKILL")
})

function runJob(args: readonly string[], timeoutMs: number): Promise<SheetJobResult> {
	return new Promise(resolve => {
		const child = spawn(process.execPath, ["--import", "tsx", RUNNER, ...args], { cwd: PACKAGE, stdio: ["ignore", "pipe", "pipe"] })
		const lines: string[] = []
		let pending = ""
		let stderr = ""
		let timedOut = false
		const timer = setTimeout(() => {
			timedOut = true
			child.kill("SIGKILL")
		}, timeoutMs)

		running.add(child)
		child.stdout.setEncoding("utf8")
		child.stdout.on("data", (chunk: string) => {
			pending += chunk

			const parts = pending.split("\n")

			pending = parts.pop() ?? ""

			for (const line of parts) if (line.startsWith("{")) lines.push(line)
		})
		child.stderr.setEncoding("utf8")
		child.stderr.on("data", (chunk: string) => {
			stderr = (stderr + chunk).slice(-4000)
		})
		child.on("close", code => {
			clearTimeout(timer)
			running.delete(child)

			if (pending.startsWith("{")) lines.push(pending)

			resolve({ args, lines, code, timedOut, stderr })
		})
	})
}

// Runs each job (arguments to this file) in its own process, `parallel` at a time.
export async function runSheetJobs(
	jobs: readonly (readonly string[])[],
	parallel: number,
	timeoutMs: number,
	onResult: (result: SheetJobResult) => void
): Promise<void> {
	let next = 0
	const worker = async (): Promise<void> => {
		while (next < jobs.length) onResult(await runJob(jobs[next++] as readonly string[], timeoutMs))
	}

	await Promise.all(Array.from({ length: Math.min(parallel, jobs.length) }, worker))
}

async function main(args: readonly string[]): Promise<void> {
	const [name = "", first = "1", count = "1", mode = "sync"] = args
	const config = sheetFuzzConfig({ ...process.env, FILEN_OFFICE_SHEET_FUZZ: process.env["FILEN_OFFICE_SHEET_FUZZ"] || "1" })

	if (config === undefined || (mode !== "sync" && mode !== "race"))
		throw new Error("usage: run.ts <file|fixture:name> <seed> <count> <sync|race>")

	const { bytes, kind, label } = target(name)
	const options = sequenceOptions(config)

	for (let seed = Number(first); seed < Number(first) + Number(count); seed++) {
		const result = await runSequence(bytes, kind, label, seed, mode, options)

		process.stdout.write(`${JSON.stringify(result)}\n`)

		// A view-only or unopenable workbook is skipped the same way by every seed.
		if (result.skipped !== undefined) break
	}
}

if (process.argv[1] === RUNNER) await main(process.argv.slice(2))
