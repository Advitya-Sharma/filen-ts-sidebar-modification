import { readFileSync } from "node:fs"
import { fileURLToPath } from "node:url"
import type { Locator, Page } from "@playwright/test"
import { test, expect } from "./fixtures"
import { LAZY_VIEWER_TIMEOUT_MS, LIVE_WRITE_TIMEOUT_MS, uploadFiles, withScratchDirectory } from "./helpers/listing"
import { DOCX_MIME } from "./helpers/docxFixtures"
import { cdpMemory, pageTimingsScript, readPageTimings, type MemorySample, type PageTimings } from "./helpers/cdpMemory"
import { resolveModKey } from "./helpers/modkey"
import { preview } from "@/locales/en/preview"

// A long editing session in the real editor (on request: FILEN_WEB_E2E_SOAK=1; FILEN_WEB_E2E_SOAK_KEYS sets
// the keys, 5,000 by default): keys typed through the page (text, Backspace, Enter) at places spread over the
// 300-page document, undo/redo bursts, select-all bold undone, and a save every 250 operations. After each
// save, on Chromium: the page's heap after a full collection, DOM nodes, layouts and the workers' heaps. The
// session holds when the heap stays flat from save to save, the DOM does not grow after the first 1,000 keys,
// no long task outside a save passes 100 ms, every blob URL made is revoked, and keys do not get slower.
// Firefox and WebKit have no heap API: keys only.

test.skip(process.env["FILEN_WEB_E2E_SOAK"] !== "1", "The soak runs on request (FILEN_WEB_E2E_SOAK=1)")
test.describe.configure({ mode: "default" })

const FIXTURE = fileURLToPath(new URL("../../filen-office/src/tests/fixtures/08-large-300p.docx", import.meta.url))
const MARKER = "Chapter 1"
const KEYS = Number(process.env["FILEN_WEB_E2E_SOAK_KEYS"] ?? "5000")
const SAVE_EVERY = 250
const PER_SPOT = 25

function percentile(values: readonly number[], share: number): number {
	const sorted = [...values].sort((a, b) => a - b)

	return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * share))] ?? 0
}

// Heap slope over `values` (MB per save), least squares.
function slope(values: readonly number[]): number {
	const n = values.length
	const meanX = (n - 1) / 2
	const meanY = values.reduce((sum, value) => sum + value, 0) / Math.max(1, n)
	let num = 0
	let den = 0

	values.forEach((value, k) => {
		num += (k - meanX) * (value - meanY)
		den += (k - meanX) ** 2
	})

	return den === 0 ? 0 : num / den
}

// The caret into a line of text in the middle of the window, after scrolling the pages by `turn` steps.
async function moveCaret(page: Page, turn: number): Promise<void> {
	await page.mouse.move(640, 400)
	await page.mouse.wheel(0, turn % 7 === 0 ? -2400 : 900)

	const runs = page.locator(".fo-docx-page .fo-docx-run")
	const box = await page.evaluate(() => {
		const middle = window.innerHeight / 2
		let best: { x: number; y: number } | null = null
		let distance = Number.POSITIVE_INFINITY

		for (const run of document.querySelectorAll(".fo-docx-page .fo-docx-run")) {
			const rect = run.getBoundingClientRect()

			if (rect.width < 20 || rect.bottom < 80 || rect.top > window.innerHeight - 40) continue

			const away = Math.abs(rect.top + rect.height / 2 - middle)

			if (away < distance) {
				distance = away
				best = { x: rect.left + Math.min(rect.width - 4, 12 + ((rect.width * 0.37) % 40)), y: rect.top + rect.height / 2 }
			}
		}

		return best
	})

	if (box === null) {
		await runs.first().click()

		return
	}

	await page.mouse.click(box.x, box.y)
}

async function save(page: Page): Promise<{ start: number; end: number }> {
	const button = page.getByRole("dialog").getByRole("button", { name: preview.previewSaveAction, exact: true })
	const start = await page.evaluate(() => performance.now())

	await button.click()
	await expect(button).toHaveCount(0, { timeout: LIVE_WRITE_TIMEOUT_MS })

	return { start, end: await page.evaluate(() => performance.now()) }
}

async function openDocument(page: Page, listbox: Locator, name: string): Promise<void> {
	await listbox.getByRole("option", { name }).dblclick()
	await expect(page.locator(".fo-docx-page").getByText(MARKER).first()).toBeVisible({ timeout: LAZY_VIEWER_TIMEOUT_MS })
	await expect(page.getByRole("button", { name: preview.previewDocxBold, exact: true })).toBeEnabled({ timeout: LAZY_VIEWER_TIMEOUT_MS })
}

test("soak: a long editing session on the 300-page document stays flat", async ({ page, browserName }) => {
	// About 40 ms a key through the page and a few seconds a save, with the uploads.
	test.setTimeout(KEYS * 120 + 30 * 60_000)
	await page.addInitScript(pageTimingsScript)

	await withScratchDirectory(page, "docx-soak", async ({ listbox, runId }) => {
		const name = `e2e-${runId}-soak-300p.docx`

		await uploadFiles(page, [{ name, mimeType: DOCX_MIME, buffer: readFileSync(FIXTURE) }], listbox)
		await openDocument(page, listbox, name)

		const mod = await resolveModKey(page)
		const memory = browserName === "chromium" ? await cdpMemory(page) : undefined
		const samples: { operations: number; keys: number; memory: MemorySample | undefined }[] = []
		const saves: { start: number; end: number }[] = []
		let operations = 0
		let typed = 0

		const operation = async (): Promise<void> => {
			operations++

			if (operations % SAVE_EVERY === 0) {
				saves.push(await save(page))
				samples.push({ operations, keys: typed, memory: await memory?.sample() })
			}
		}

		await moveCaret(page, 0)

		try {
			while (typed < KEYS) {
				if (typed % PER_SPOT === 0) await moveCaret(page, typed / PER_SPOT)

				const kind = typed % 20

				if (kind === 7 || kind === 16) await page.keyboard.press("Backspace")
				else if (kind === 9 || kind === 17) await page.keyboard.press("Enter")
				else await page.keyboard.type("a")

				typed++
				await operation()

				// An undo/redo burst every 25 keys, a select-all bold taken back every 250.
				if (typed % 25 === 12) {
					for (let k = 0; k < 3; k++) await page.keyboard.press(`${mod}+z`)
					for (let k = 0; k < 3; k++) await page.keyboard.press(`${mod}+Shift+z`)
					await operation()
				}

				if (typed % 250 === 125) {
					await page.keyboard.press(`${mod}+a`)
					await page.keyboard.press(`${mod}+b`)
					await page.keyboard.press(`${mod}+z`)
					await moveCaret(page, typed)
					await operation()
				}
			}

			saves.push(await save(page))
			samples.push({ operations, keys: typed, memory: await memory?.sample() })
		} finally {
			await memory?.dispose()
		}

		const timings: PageTimings = await readPageTimings(page)
		const decile = Math.max(1, Math.floor(timings.keys.length / 10))
		const firstP95 = percentile(timings.keys.slice(0, decile), 0.95)
		const lastP95 = percentile(timings.keys.slice(-decile), 0.95)
		const outside = timings.longTasks.filter(
			task => !saves.some(window => task.start + task.duration >= window.start && task.start <= window.end)
		)

		for (const sample of samples) {
			const heap = sample.memory

			console.log(
				`after ${String(sample.operations)} operations (${String(sample.keys)} keys)` +
					(heap === undefined
						? ""
						: `: heap ${heap.jsHeapUsedMb.toFixed(1)} MB, ${String(heap.domNodes)} DOM nodes, ${String(heap.listeners)} listeners, ` +
							`${String(heap.layoutCount)} layouts, workers ${heap.workers.map(worker => `${worker.usedMb.toFixed(1)} MB`).join(", ") || "none"}`)
			)
		}

		console.log(
			`${browserName}: ${String(timings.keys.length)} keys, key to frame p95 ${firstP95.toFixed(1)} ms first decile, ${lastP95.toFixed(1)} ms last; ` +
				`${String(outside.length)} long tasks outside saves (longest ${Math.max(0, ...outside.map(task => task.duration)).toFixed(0)} ms); ` +
				`blob URLs ${String(timings.created)} made, ${String(timings.revoked)} revoked`
		)

		await page.keyboard.press("Escape")

		// Firefox and WebKit time keys to the millisecond: their drift is what they can tell.
		expect(lastP95).toBeLessThanOrEqual(Math.max(firstP95 * 1.2, firstP95 + 2))

		if (memory === undefined) {
			return
		}

		const heaps = samples.flatMap(sample => (sample.memory === undefined ? [] : [sample.memory]))
		const afterThousand = samples.find(sample => sample.keys >= 1000)?.memory?.domNodes
		const lastNodes = heaps[heaps.length - 1]?.domNodes ?? 0

		expect(slope(heaps.slice(-15).map(heap => heap.jsHeapUsedMb))).toBeLessThan(1)

		// A shorter run (FILEN_WEB_E2E_SOAK_KEYS) has no count after 1,000 keys to hold to.
		if (afterThousand !== undefined) {
			expect(lastNodes).toBeLessThanOrEqual(afterThousand * 1.05)
		}
		expect(outside.filter(task => task.duration > 100)).toEqual([])
		expect(timings.revoked).toBeGreaterThanOrEqual(timings.created)
	})
})
