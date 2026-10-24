import { PerformanceObserver } from "node:perf_hooks"
import { describe, expect, it } from "vitest"
import { planRefEdit } from "../core/xlsx/refs/index"
import { saveXlsx } from "../core/xlsx/write/save"
import { buildFixture, openWorkbook } from "./xlsxWriteFixtures"
import { budget } from "./timing"

// A save after a column insert on a million cells, on an in-thread host: every row spliced and checked (the store
// and the rows read back) in slices, so the thread is never held over 50 ms of its CPU time.

const ROWS = 100_000

function body(): string {
	let data = "<sheetData>"

	for (let r = 1; r <= ROWS; r++) {
		data += `<row r="${r}" spans="1:10">`

		for (let c = 0; c < 10; c++) data += `<c r="${String.fromCharCode(65 + c)}${r}"><v>${(r * 10 + c) / 7}</v></c>`

		data += "</row>"
	}

	return `${data}</sheetData>`
}

describe("a sliced save after a column insert", () => {
	it("gives the host a turn at least every 50 ms on 100,000 rows of 10 cells", { timeout: 300_000 }, async () => {
		const workbook = await openWorkbook(buildFixture({ sheets: [{ name: "Data", body: body() }] }))

		await workbook.load(0)

		const planned = await planRefEdit(workbook, { kind: "axis", sheet: 0, shift: { type: "insert", axis: "cols", at: 0, count: 1 } })

		if (!planned.ok) throw new Error(JSON.stringify(planned.refusal))

		planned.plan.apply()

		// The longest the thread was held between two turns, in this thread's CPU time (what the host's other work
		// waits for; the wall clock of a loaded machine would count other processes' time too). The slices follow
		// the same clock.
		const threadMs = (): number => {
			const usage = process.threadCpuUsage()

			return (usage.user + usage.system) / 1000
		}
		// Garbage collection pauses (of what the fixture left as much as of the save) are the runtime's, not the
		// save's slices: their time within a stretch is left out.
		const collections: { start: number; end: number }[] = []
		const observer = new PerformanceObserver(list => {
			for (const entry of list.getEntries()) collections.push({ start: entry.startTime, end: entry.startTime + entry.duration })
		})

		observer.observe({ entryTypes: ["gc"] })

		let longest = 0
		let resumed = threadMs()
		let resumedAt = performance.now()
		let turns = 0
		const stretch = (): number => {
			const now = performance.now()
			let collecting = 0

			for (const pause of collections) collecting += Math.max(0, Math.min(now, pause.end) - Math.max(resumedAt, pause.start))

			return threadMs() - resumed - collecting
		}
		const slicing = {
			sliceMs: 8,
			now: threadMs,
			pause: async () => {
				longest = Math.max(longest, stretch())
				turns++
				await new Promise(resolve => setTimeout(resolve, 0))
				resumed = threadMs()
				resumedAt = performance.now()
			}
		}
		const cpu = threadMs()
		const started = performance.now()
		const result = await saveXlsx({ opc: workbook.opc, workbook }, { slicing })
		const wall = performance.now() - started
		const used = threadMs() - cpu

		// GC entries are delivered asynchronously.
		await new Promise(resolve => setTimeout(resolve, 10))
		longest = Math.max(longest, stretch())
		observer.disconnect()

		expect(result.ok, JSON.stringify(result)).toBe(true)
		console.log(
			`sliced column-insert save, ${ROWS} rows: ${wall.toFixed(0)} ms wall, ${used.toFixed(0)} ms CPU, ${turns} turns, longest task ${longest.toFixed(1)} ms, ${JSON.stringify(result.ok ? result.timings : null)}`
		)
		expect(longest).toBeLessThan(budget(50))
	})
})
