import { PerformanceObserver } from "node:perf_hooks"
import { describe, expect, it } from "vitest"
import { planRefEdit } from "../core/xlsx/refs/index"
import { saveXlsx } from "../core/xlsx/write/save"
import { buildFixture, openWorkbook } from "./xlsxWriteFixtures"

// A save after a structure edit on a large sheet, on an in-thread host: written and checked in slices, so the
// thread is never held over 50 ms of its CPU time, and its cost (wall and this thread's CPU time) is logged.

const ROWS = 100_000

function body(): string {
	let data = "<sheetData>"

	for (let r = 1; r <= ROWS; r++) {
		data += `<row r="${r}"><c r="A${r}"><v>${r}</v></c><c r="B${r}"><v>${r * 2}</v></c><c r="C${r}"><f>A${r}+B${r}</f><v>0</v></c>`
		data +=
			r === 1
				? `<c r="D1"><f t="shared" ref="D1:D${ROWS}" si="0">C1*2</f><v>0</v></c>`
				: `<c r="D${r}"><f t="shared" si="0"/><v>0</v></c>`
		data += "</row>"
	}

	return `${data}</sheetData>`
}

describe("a sliced save", () => {
	it("gives the host a turn at least every 50 ms after a structure edit on 100,000 rows of formulas", { timeout: 300_000 }, async () => {
		const workbook = await openWorkbook(buildFixture({ sheets: [{ name: "Data", body: body() }] }))

		await workbook.load(0)

		const planned = await planRefEdit(workbook, { kind: "axis", sheet: 0, shift: { type: "insert", axis: "rows", at: 1, count: 1 } })

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
			`sliced save, ${ROWS} rows: ${wall.toFixed(0)} ms wall, ${used.toFixed(0)} ms CPU, ${turns} turns, longest task ${longest.toFixed(1)} ms, ${JSON.stringify(result.ok ? result.timings : null)}`
		)
		expect(longest).toBeLessThan(50)
	})
})
