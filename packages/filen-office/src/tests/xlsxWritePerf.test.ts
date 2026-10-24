import { zipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { encodeUtf8 } from "../core/xml/index"
import { nativeDeflate, nativeDeflateChunks } from "../core/docx/write/zip"
import { CellKind } from "../core/xlsx/model/types"
import { saveXlsx } from "../core/xlsx/write/save"
import { bytesSource } from "../core/opc/index"
import { openXlsx } from "../core/xlsx/read/index"
import { R, SML } from "./xlsxWriteFixtures"
import { ZIP_ENTRY_TIME } from "./zipTime"
import { budget } from "./timing"

// The save budget at 1M cells: after a one-cell edit, ≤ 0.6 s of CPU on the thread that saves, only the edited
// part deflated (every other entry copied as stored), and a peak of at most the open's memory plus twice the
// edited part. Opt-in (FILEN_OFFICE_XLSX_PERF=1): the workbook is built in the test.

const PERF = process.env["FILEN_OFFICE_XLSX_PERF"] === "1"
const ROWS = 100_000
const COLS = 10

function bigWorkbook(): { bytes: Uint8Array; sheetBytes: number } {
	const chunks: string[] = []

	chunks.push(
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${SML}" xmlns:r="${R}"><dimension ref="A1:J${ROWS}"/><sheetData>`
	)

	for (let row = 1; row <= ROWS; row++) {
		let markup = `<row r="${row}" spans="1:${COLS}">`

		for (let col = 0; col < COLS; col++) {
			const reference = `${String.fromCharCode(65 + col)}${row}`

			markup +=
				col === 0 ? `<c r="${reference}" t="s"><v>${row % 100}</v></c>` : `<c r="${reference}"><v>${(row * COLS + col) / 7}</v></c>`
		}

		chunks.push(`${markup}</row>`)
	}

	chunks.push("</sheetData></worksheet>")

	const sheet = encodeUtf8(chunks.join(""))
	const strings = Array.from({ length: 100 }, (_, k) => `<si><t>text ${k}</t></si>`).join("")
	const files: Record<string, Uint8Array> = {
		"[Content_Types].xml": encodeUtf8(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
				'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>'
		),
		"_rels/.rels": encodeUtf8(
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`
		),
		"xl/workbook.xml": encodeUtf8(
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="${SML}" xmlns:r="${R}"><sheets><sheet name="Big" sheetId="1" r:id="rId1"/></sheets></workbook>`
		),
		"xl/_rels/workbook.xml.rels": encodeUtf8(
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`
		),
		"xl/sharedStrings.xml": encodeUtf8(
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${SML}" count="${ROWS}" uniqueCount="100">${strings}</sst>`
		),
		"xl/worksheets/sheet1.xml": sheet
	}

	return { bytes: zipSync(files, { level: 6, mtime: ZIP_ENTRY_TIME }), sheetBytes: sheet.length }
}

function memory(): number {
	const usage = process.memoryUsage()

	return usage.heapUsed + usage.arrayBuffers
}

describe.skipIf(!PERF)("xlsx save at 1M cells", () => {
	it("saves a one-cell edit within budget", { timeout: 300_000 }, async () => {
		const { bytes, sheetBytes } = bigWorkbook()
		// Inflated natively, as the worker does.
		const inflate = (input: AsyncIterable<Uint8Array>): AsyncIterable<Uint8Array> =>
			ReadableStream.from(input).pipeThrough(new DecompressionStream("deflate-raw"))
		const opened = await openXlsx(bytesSource(bytes), { inflate })

		if (!opened.ok) throw new Error(opened.message)

		const workbook = opened.workbook
		const sheet = await workbook.load(0)

		expect(sheet.store.cellCount).toBe(ROWS * COLS)

		sheet.store.setCell(ROWS / 2, 3, CellKind.NUMBER, 0.5, 0, -1)
		workbook.changes.sheetRows(0)

		const native = nativeDeflate()
		const nativeChunks = nativeDeflateChunks()
		const deflated: number[] = []
		const zip = {
			...(native === undefined ? {} : { deflate: async (data: Uint8Array) => (deflated.push(data.length), native(data)) }),
			...(nativeChunks === undefined
				? {}
				: {
						deflateChunks: async (chunks: AsyncIterable<Uint8Array>) => {
							let length = 0
							const counted = async function* (): AsyncIterable<Uint8Array> {
								for await (const chunk of chunks) {
									length += chunk.length
									yield chunk
								}
							}
							const out = await nativeChunks(counted())

							deflated.push(length)

							return out
						}
					})
		}
		const before = memory()
		let peak = before
		const sampler = setInterval(() => {
			peak = Math.max(peak, memory())
		}, 5)
		// This thread's CPU, as the baseline measured it (zlib's native inflate and deflate run off it), and the
		// whole process's.
		const thread = process.threadCpuUsage()
		const all = process.cpuUsage()
		const result = await saveXlsx({ opc: workbook.opc, workbook }, { zip })
		const threadUsed = process.threadCpuUsage(thread)
		const allUsed = process.cpuUsage(all)

		clearInterval(sampler)
		peak = Math.max(peak, memory())

		const cpuMs = (threadUsed.user + threadUsed.system) / 1000
		const processMs = (allUsed.user + allUsed.system) / 1000

		console.log(
			`xlsx save 1M cells: ${cpuMs.toFixed(0)} ms thread CPU, ${processMs.toFixed(0)} ms process CPU (${JSON.stringify(result.ok ? result.timings : result)}); deflated ${deflated.join(", ")} bytes; sheet ${sheetBytes} bytes; peak +${((peak - before) / 1048576).toFixed(1)} MB`
		)

		expect(result.ok).toBe(true)
		// The edited sheet alone (its new size is within a few bytes of the source's).
		expect(deflated.length).toBe(1)
		expect(Math.abs((deflated[0] ?? 0) - sheetBytes)).toBeLessThan(64)
		expect(cpuMs).toBeLessThanOrEqual(budget(600))
		expect(peak - before).toBeLessThanOrEqual(2 * sheetBytes)
	})
})
