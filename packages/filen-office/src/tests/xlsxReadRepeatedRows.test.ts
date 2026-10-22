import { zipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { bytesSource } from "../core/opc/index"
import { openXlsx } from "../core/xlsx/read/index"
import { encodeUtf8 } from "../core/xml/index"
import { SML, workbookFiles } from "./xlsxReadFixtures"
import { cpuMs } from "./timing"
import { ZIP_ENTRY_TIME } from "./zipTime"

// Repeated rows are merged in steps on the reader's slice clock: no long task, and an abort is heard meanwhile.

function repeated(): Uint8Array {
	const row = `<row r="1">${"<c><v>1</v></c>".repeat(16_384)}</row>`
	const sheet = encodeUtf8(`<worksheet xmlns="${SML}"><sheetData>${row.repeat(30)}</sheetData></worksheet>`)

	return zipSync(workbookFiles({ sheets: [{ bytes: sheet }] }), { mtime: ZIP_ENTRY_TIME })
}

describe("xlsx reader: repeated rows", () => {
	it("merges 30 copies of a 16,384-cell row without one long synchronous stretch", async () => {
		const opened = await openXlsx(bytesSource(repeated()))

		if (!opened.ok) throw new Error(opened.message)

		let longest = 0
		let last = cpuMs()
		const slicing = {
			sliceMs: 8,
			now: () => performance.now(),
			pause: () => {
				longest = Math.max(longest, cpuMs() - last)

				return new Promise<void>(resolve =>
					setImmediate(() => {
						last = cpuMs()
						resolve()
					})
				)
			}
		}

		last = cpuMs()

		const model = await opened.workbook.load(0, slicing)

		longest = Math.max(longest, cpuMs() - last)
		expect(model.store.cellCount).toBe(16_384)
		expect(longest).toBeLessThan(50)
	})
})
