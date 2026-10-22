import { inThreadSheetHost } from "../dom/sheet/host"
import { buildXlsx } from "../tests/oracles/sheet/synthetic"

// Dynamic arrays spilling and reshaping in the in-thread host with its default slicing: per step, CPU time of this
// thread, wall time and the longest task (the most time between two turns of the event loop), which the budget holds
// under 50 ms (the engine's first build aside).
//   node --import tsx src/bench/sheetSpillBench.ts [--json]

const STEPS: readonly (readonly [string, string])[] = [
	["spill 150k", "=SEQUENCE(150000)"],
	["reshape 150k to 20k x 5", "=SEQUENCE(20000,5)"],
	["reshape to 20k x 1", "=SEQUENCE(20000)"],
	["reshape 20k x 1 to 4k x 5", "=SEQUENCE(4000,5)"]
]

const bytes = buildXlsx({
	sheets: [{ name: "S", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' }],
	application: { name: "Microsoft Excel", version: "16.0300" },
	calcPr: 'calcId="191029"'
})
const opened = await inThreadSheetHost().open(bytes, { kind: "xlsx" }, () => undefined)

if (opened.kind !== "opened") throw new Error(`open: ${opened.kind}`)

const { session, info } = opened
const sheet = info.activeSheet

await session.tiles(sheet, Uint32Array.of(0, 0))
await session.intent()
// The engine built by a first edit, apart from the steps.
await session.command({ type: "cells.set", sheet, cells: { rows: Uint32Array.of(0), cols: Uint16Array.of(1), inputs: ["=A1"] } })

let last = performance.now()
let longest = 0
let running = true
const tick = (): void => {
	const now = performance.now()

	longest = Math.max(longest, now - last)
	last = now

	if (running) setImmediate(tick)
}

setImmediate(tick)

const results: Record<string, unknown>[] = []

for (const [name, input] of STEPS) {
	await new Promise(resolve => setImmediate(resolve))
	longest = 0
	last = performance.now()

	const cpu = process.threadCpuUsage()
	const started = performance.now()
	const result = await session.command({
		type: "cells.set",
		sheet,
		cells: { rows: Uint32Array.of(0), cols: Uint16Array.of(2), inputs: [input] }
	})
	const used = process.threadCpuUsage(cpu)

	if (!result.ok) throw new Error(`${name}: refused ${result.refusal.reason}`)

	results.push({
		step: name,
		cpuMs: Math.round((used.user + used.system) / 1000),
		wallMs: Math.round(performance.now() - started),
		longestTaskMs: Math.round(longest)
	})
}

// Structure edits through a 150k spill.
await session.command({
	type: "cells.set",
	sheet,
	cells: { rows: Uint32Array.of(0), cols: Uint16Array.of(2), inputs: ["=SEQUENCE(150000)"] }
})

for (const [name, command] of [
	["insert row through a 150k spill", { type: "axis.insert", sheet, axis: "rows", at: 9, count: 1 }],
	["delete row through a 150k spill", { type: "axis.delete", sheet, axis: "rows", at: 9, count: 1 }],
	["insert column through a 150k spill", { type: "axis.insert", sheet, axis: "cols", at: 0, count: 1 }],
	["delete column through a 150k spill", { type: "axis.delete", sheet, axis: "cols", at: 0, count: 1 }]
] as const) {
	await new Promise(resolve => setImmediate(resolve))
	longest = 0
	last = performance.now()

	const cpu = process.threadCpuUsage()
	const started = performance.now()
	const result = await session.command(command)
	const used = process.threadCpuUsage(cpu)

	if (!result.ok) throw new Error(`${name}: refused ${result.refusal.reason}`)

	results.push({
		step: name,
		cpuMs: Math.round((used.user + used.system) / 1000),
		wallMs: Math.round(performance.now() - started),
		longestTaskMs: Math.round(longest)
	})
}

running = false
session.close()

if (process.argv.includes("--json")) process.stdout.write(`${JSON.stringify(results)}\n`)
else console.table(results)
