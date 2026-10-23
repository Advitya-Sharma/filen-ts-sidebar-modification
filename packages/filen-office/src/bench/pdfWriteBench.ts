import { PdfWriter } from "../core/pdfwrite/index"
import { buildSyntheticPdf, JPEG_32x24, StubFont } from "../tests/pdfWriteFixtures"

// PDF writer cost for 300 pages: CPU per page (process time of the writer, glyph runs prepared in advance)
// and peak memory (RSS and ArrayBuffers sampled at every page end, over a baseline taken after a GC).
// Two documents: text pages like a DOCX export (48 tagged lines of ~95 characters, a rule, a picture) and the
// all-features synthetic page.
//
//   node --expose-gc --import tsx src/bench/pdfWriteBench.ts

const PAGES = 300
const LINE = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore."

function gc(): void {
	const collect = (globalThis as { gc?: () => void }).gc

	collect?.()
}

function memory(): { rss: number; buffers: number; heap: number } {
	const usage = process.memoryUsage()

	return { rss: usage.rss, buffers: usage.arrayBuffers, heap: usage.heapUsed }
}

async function textDocument(
	password: string | undefined
): Promise<{ cpuMs: number; bytes: number; rssPeak: number; buffersPeak: number; heapPeak: number; retainedPeak: number }> {
	const font = new StubFont()
	const runs = Array.from({ length: 48 }, (_, line) => font.shape(`${line + 1}. ${LINE}`))

	gc()

	const base = memory()
	let rssPeak = 0
	let buffersPeak = 0
	let heapPeak = 0
	let retainedPeak = 0
	const start = process.cpuUsage()
	let resumed = performance.now()
	const writer = await PdfWriter.create({
		title: "Bench",
		lang: "en",
		tagged: true,
		...(password === undefined ? {} : { password }),
		slicing: {
			sliceMs: 16,
			now: () => performance.now(),
			pause: () => {
				longestSlice = Math.max(longestSlice, performance.now() - resumed)

				return new Promise<void>(resolve =>
					setImmediate(() => {
						resumed = performance.now()
						resolve()
					})
				)
			}
		}
	})
	const picture = await writer.image({ data: JPEG_32x24, key: "image1" })

	// Key derivation waits on WebCrypto (off the thread): slices start with the pages.
	resumed = performance.now()

	for (let index = 0; index < PAGES; index++) {
		const page = writer.beginPage(816, 1056)

		page.beginArtifact({ type: "Pagination", subtype: "Header" })
		page.text([font.shape(`Page ${index + 1}`)], { x: 700, y: 48, size: 12 })
		page.endArtifact()

		runs.forEach((run, line) => {
			const paragraph = writer.open("P")

			page.beginContent(paragraph)
			page.text([run], { x: 96, y: 120 + line * 18, size: 14.67 })
			page.endContent()
			writer.close(paragraph)
		})

		page.beginArtifact({ type: "Layout" })
		page.strokeColor(0x808080ff).lineWidth(1).moveTo(96, 1000).lineTo(720, 1000).stroke()
		page.endArtifact()

		const figure = writer.open("Figure", { alt: "Picture" })

		page.beginContent(figure)
		page.drawImage(picture, 600, 900, 96, 72)
		page.endContent()
		writer.close(figure)
		writer.outline(`Page ${index + 1}`, 1, { page: index, y: 120 })
		await writer.endPage(page)

		// Every 50 pages: what stays reachable (garbage collected first).
		if (index % 50 === 49) {
			const collecting = performance.now()

			gc()
			// The forced collection is the bench's, not the writer's: keep it out of the slice.
			resumed += performance.now() - collecting
			retainedPeak = Math.max(retainedPeak, memory().heap + memory().buffers - base.heap - base.buffers)
		}

		const now = memory()

		rssPeak = Math.max(rssPeak, now.rss - base.rss)
		buffersPeak = Math.max(buffersPeak, now.buffers - base.buffers)
		heapPeak = Math.max(heapPeak, now.heap - base.heap)
	}

	const finishStart = performance.now()
	const output = await writer.finish()
	const cpu = process.cpuUsage(start)

	finishMs = performance.now() - finishStart
	const now = memory()

	rssPeak = Math.max(rssPeak, now.rss - base.rss)
	buffersPeak = Math.max(buffersPeak, now.buffers - base.buffers)

	heapPeak = Math.max(heapPeak, now.heap - base.heap)

	return { cpuMs: (cpu.user + cpu.system) / 1000, bytes: output.size, rssPeak, buffersPeak, heapPeak, retainedPeak }
}

let finishMs = 0
let longestSlice = 0
const mb = (bytes: number): string => `${(bytes / 1024 / 1024).toFixed(1)} MB`

for (const password of [undefined, "secret"]) {
	// One warm-up run (JIT), then the measured one.
	await textDocument(password)

	longestSlice = 0

	const result = await textDocument(password)

	console.log(
		`text ${PAGES} pages${password === undefined ? "" : " encrypted"}: ${result.cpuMs.toFixed(0)} ms CPU, ${(result.cpuMs / PAGES).toFixed(2)} ms/page, ` +
			`output ${mb(result.bytes)}, peak RSS +${mb(result.rssPeak)}, peak ArrayBuffers +${mb(result.buffersPeak)}, peak JS heap +${mb(result.heapPeak)} (retained after GC +${mb(result.retainedPeak)}), finish ${finishMs.toFixed(1)} ms, longest slice ${longestSlice.toFixed(1)} ms`
	)
}

{
	await buildSyntheticPdf({ pages: 10, deterministic: false })
	gc()

	const base = memory()
	const start = process.cpuUsage()
	const { bytes } = await buildSyntheticPdf({ pages: PAGES, deterministic: false })
	const cpu = process.cpuUsage(start)
	const now = memory()

	console.log(
		`synthetic ${PAGES} pages: ${((cpu.user + cpu.system) / 1000).toFixed(0)} ms CPU, ${((cpu.user + cpu.system) / 1000 / PAGES).toFixed(2)} ms/page, ` +
			`output ${mb(bytes.length)}, RSS +${mb(now.rss - base.rss)} at the end (output held twice: parts and the joined copy)`
	)
}
