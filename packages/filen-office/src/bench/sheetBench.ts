import { readFileSync } from "node:fs"
import { basename } from "node:path"
import v8 from "node:v8"
import { bytesSource } from "@filen/office/core/opc"
import { openXlsxSession } from "../core/sheet/commands"
import { CsvDocument } from "../core/sheet/csv/csvDocument"
import { parseCsvFile } from "../core/sheet/csv/csvView"
import type { SheetChange, SheetCommandResult, SheetEvent } from "../core/sheet/protocol"
import type { SheetSession } from "../core/sheet/session"
import type { XlsxDocument } from "../core/xlsx/edit/document"
import { saveXlsx } from "../core/xlsx/write/index"

// The spreadsheet engine's budgets, workload by workload, in one Node process (the worker's thread: the in-thread
// host runs the same session code): CPU time of this thread, wall time, retained heap after a full collection,
// the process's peak RSS, and the size of what would cross to the page (v8.serialize, postMessage's serializer).
// Compared against the previous engine's numbers at the same fixtures (HEAD in the baseline document).
//   node --expose-gc --import tsx src/bench/sheetBench.ts <fixture.xlsx|.csv> [--keys=50] [--json]
// The fixtures: 100k×10 values (m0), with a formula column and no stored results (m1), with stored results (m2),
// 300k×10 values, and the same 100k×10 as CSV.

const gc = (globalThis as { gc?: () => void }).gc

if (gc === undefined) throw new Error("run with --expose-gc")

const MB = 1024 * 1024
const args = process.argv.slice(2)
const file = args.find(arg => !arg.startsWith("--"))

if (file === undefined) throw new Error("usage: sheetBench.ts <file> [--keys=50] [--json]")

const option = (name: string, fallback: string): string =>
	args.find(arg => arg.startsWith(`--${name}=`))?.slice(name.length + 3) ?? fallback
const KEYS = Number(option("keys", "50"))
const IMMEDIATE = { sliceMs: 1_000_000, now: () => 0, pause: () => Promise.resolve() }

// The previous engine's numbers (HEAD worker path, Node, median of 3): CPU ms and MB.
const HEAD: Readonly<Record<string, Readonly<Record<string, number>>>> = {
	m0: {
		openCpuMs: 1232,
		openRetainedMb: 23.5,
		firstEditCpuMs: 0.9,
		keyP95CpuMs: 0.04,
		insertRowCpuMs: 245,
		insertRowMessageMb: 28.6,
		saveCpuMs: 327
	},
	m1: {
		openCpuMs: 2607,
		openRetainedMb: 249.1,
		firstEditCpuMs: 1.7,
		keyP95CpuMs: 0.12,
		insertRowCpuMs: 526,
		insertRowMessageMb: 31.3,
		saveCpuMs: 467
	},
	m2: {
		openCpuMs: 1497,
		openRetainedMb: 76.7,
		firstEditCpuMs: 1192,
		keyP95CpuMs: 0.13,
		insertRowCpuMs: 474,
		insertRowMessageMb: 31.5,
		saveCpuMs: 456
	},
	"3m": { openCpuMs: 4093, openRetainedMb: 69.4, openPeakRssMb: 650 },
	csv: { openCpuMs: 372, openRetainedMb: 36.6 }
}

function retained(): number {
	gc?.()
	gc?.()

	const usage = process.memoryUsage()

	return (usage.heapUsed + usage.external) / MB
}

function peakRss(): number {
	return process.resourceUsage().maxRSS / 1024
}

function cpuMs(since: NodeJS.CpuUsage): number {
	const used = process.threadCpuUsage(since)

	return (used.user + used.system) / 1000
}

function round(value: number, digits = 1): number {
	const scale = 10 ** digits

	return Math.round(value * scale) / scale
}

function percentile(values: readonly number[], p: number): number {
	const sorted = [...values].sort((a, z) => a - z)

	return sorted[Math.min(sorted.length - 1, Math.floor((p / 100) * sorted.length))] ?? 0
}

function changeOf(result: SheetCommandResult): SheetChange {
	if (!result.ok) throw new Error(`refused: ${result.refusal.reason} ${result.refusal.detail ?? ""}`)

	return result.change
}

const results: Record<string, unknown> = { file: basename(file), node: process.version }

async function measure<T>(
	name: string,
	run: () => Promise<T> | T,
	extra?: (value: T) => Record<string, unknown>
): Promise<{ value: T; cpu: number }> {
	const before = retained()
	const cpu = process.threadCpuUsage()
	const started = performance.now()
	const value = await run()
	const cpuUsed = cpuMs(cpu)
	const wall = performance.now() - started
	const after = retained()

	results[name] = {
		cpuMs: round(cpuUsed),
		wallMs: round(wall),
		retainedDeltaMb: round(after - before),
		peakRssMb: Math.round(peakRss()),
		...(extra?.(value) ?? {})
	}

	return { value, cpu: cpuUsed }
}

function fixtureKey(name: string): string {
	if (name.endsWith(".csv")) return "csv"
	if (name.includes("300000")) return "3m"

	return /-(m\d)\./.exec(name)?.[1] ?? name
}

// Where a save's time goes (the writer, the gate's part check, the zip, the gate's re-read of the zip), from the
// splice writer itself on the session's model.
async function savePhases(session: SheetSession<XlsxDocument>): Promise<Record<string, number> | string> {
	const workbook = session.document.workbook
	const saved = await saveXlsx({ opc: workbook.opc, workbook })

	if (!saved.ok) return saved.reason

	return Object.fromEntries(Object.entries(saved.timings).map(([name, ms]) => [name, round(ms)]))
}

async function benchCsv(bytes: Uint8Array): Promise<void> {
	const base = retained()
	const { value: document } = await measure("open", () => {
		const { rows, format } = parseCsvFile(bytes, false)

		return { document: new CsvDocument(rows, format), rows: rows.length }
	})

	results["openRetainedMb"] = round(retained() - base)
	results["rows"] = document.rows
	results["writable"] = document.document.writable
}

async function benchXlsx(bytes: Uint8Array): Promise<void> {
	const base = retained()
	const events: SheetEvent[] = []
	const started = performance.now()
	const { value: opened } = await measure("open", () =>
		openXlsxSession(bytesSource(bytes), { kind: "xlsx", slicing: IMMEDIATE }, event => events.push(event))
	)

	if (opened.kind !== "opened") throw new Error(`open: ${opened.kind}`)

	const { session, info } = opened
	const sheet = info.activeSheet
	// The first viewport (a 1400x900 window shows about 45 rows by 10 columns: one tile) and its message.
	const { value: first } = await measure(
		"firstTile",
		() => session.tiles(sheet, Uint32Array.of(0, 0)),
		value => ({
			messageBytes: v8.serialize(value).byteLength,
			sinceOpenStartMs: round(performance.now() - started)
		})
	)

	results["openInfoBytes"] = v8.serialize(info).byteLength
	results["openRetainedMb"] = round(retained() - base)
	results["firstTileCells"] = first.tiles[0]?.count ?? 0

	if (info.activeLayout.rowCount === 0) return

	// The proof, then the engine on edit intent (it never blocks a commit: it is built after intent answers).
	await measure(
		"proof",
		() => session.intent(),
		value => ({ state: value.state })
	)

	const verdict = await session.intent()

	if (verdict.state !== "writable") {
		results["viewOnly"] = verdict.state === "viewOnly" ? verdict.reason : verdict.state
		session.close()

		return
	}

	await measure(
		"engine",
		() => session.layout(sheet),
		() => ({ events: events.filter(event => event.type === "engine").length })
	)
	results["afterEngineRetainedMb"] = round(retained() - base)

	// Keystroke commits in column B (column J's formulas depend on it in m1 and m2).
	const keyCpu: number[] = []
	const keyBytes: number[] = []

	await measure("firstEdit", async () =>
		changeOf(
			await session.command({ type: "cells.set", sheet, cells: { rows: Uint32Array.of(1), cols: Uint16Array.of(1), inputs: ["7"] } })
		)
	)

	// The budget's save: one edited cell at 1M cells (only the edited row is serialised; nothing else re-deflated).
	await measure(
		"saveOneEdit",
		() => session.save(),
		saved => (saved.ok ? { bytes: saved.bytes.byteLength, changedParts: saved.changedParts.length } : { refused: saved.reason })
	)
	results["saveOneEditPhases"] = await savePhases(session)

	for (let i = 0; i < KEYS; i++) {
		const cpu = process.threadCpuUsage()
		const change = changeOf(
			await session.command({
				type: "cells.set",
				sheet,
				cells: { rows: Uint32Array.of(2 + i * 37), cols: Uint16Array.of(1), inputs: [String(i * 3.5)] }
			})
		)

		keyCpu.push(cpuMs(cpu))
		keyBytes.push(v8.serialize(change).byteLength)
	}

	results["keys"] = {
		count: KEYS,
		cpuP50Ms: round(percentile(keyCpu, 50), 3),
		cpuP95Ms: round(percentile(keyCpu, 95), 3),
		messageP95Bytes: percentile(keyBytes, 95)
	}

	await measure(
		"insertRow",
		async () => changeOf(await session.command({ type: "axis.insert", sheet, axis: "rows", at: 10, count: 1 })),
		change => ({ messageBytes: v8.serialize(change).byteLength })
	)

	// Scrolling: tiles down the sheet, each cold.
	const tileCpu: number[] = []

	for (let row = 0; row < Math.min(info.activeLayout.rowCount, 100_000); row += 6_400) {
		const cpu = process.threadCpuUsage()

		await session.tiles(sheet, Uint32Array.of(Math.floor(row / 64), 0))
		tileCpu.push(cpuMs(cpu))
	}

	results["scrollTiles"] = {
		count: tileCpu.length,
		cpuMaxMs: round(Math.max(...tileCpu), 2),
		cpuP95Ms: round(percentile(tileCpu, 95), 2)
	}

	// After the keys and an inserted row (every row below renumbered).
	await measure(
		"saveAfterAll",
		() => session.save(),
		saved => (saved.ok ? { bytes: saved.bytes.byteLength, changedParts: saved.changedParts.length } : { refused: saved.reason })
	)
	results["saveAfterAllPhases"] = await savePhases(session)
	results["endRetainedMb"] = round(retained() - base)
	session.close()
}

const bytes = new Uint8Array(readFileSync(file))
const key = fixtureKey(basename(file))

if (file.endsWith(".csv")) await benchCsv(bytes)
else await benchXlsx(bytes)

results["peakRssMb"] = Math.round(peakRss())

const head = HEAD[key]

if (head !== undefined) results["head"] = head

if (args.includes("--json")) process.stdout.write(`${JSON.stringify(results)}\n`)
else console.log(JSON.stringify(results, null, 1))
