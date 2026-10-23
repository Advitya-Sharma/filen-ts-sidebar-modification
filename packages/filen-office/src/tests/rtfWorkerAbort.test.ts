import { describe, expect, it } from "vitest"
import type { Slicing } from "../core/slicing"
import { RtfAborted, openRtfSource, proveRtf } from "../core/rtf/write/writer"
import { serveDocxWorker } from "../dom/docx/host"
import { docxRtfServer } from "../dom/docx/rtfWorker"
import { editRtf, editsOf, latin1, paragraphs } from "./rtfWriteFixtures"

// A long RTF job in the worker (reading the file, Gate 1's proof, a save) runs in slices; the document's close ends it at the next
// slice: nothing more runs and no answer goes back.

const rtf = (count: number): Uint8Array =>
	latin1(
		"{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Times;}{\\f1 Arial;}}" +
			Array.from({ length: count }, (_, k) => `\\pard\\plain\\f${k % 2} Paragraph ${k} with {\\b bold} text.\\par\n`).join("") +
			"}"
	)

const macrotask = (): Promise<void> => new Promise(resolve => setTimeout(resolve, 0))

// Slices of no length: a pause at every step, counted; `onPause` runs before each.
function counting(onPause: (count: number) => void): { slicing: Slicing; pauses: () => number } {
	let count = 0

	return {
		slicing: {
			pause: () => {
				onPause(++count)

				return macrotask()
			},
			sliceMs: 0,
			now: () => Date.now()
		},
		pauses: () => count
	}
}

describe("aborting RTF jobs in the worker", () => {
	it("reads and proves in slices, as without them", async () => {
		const bytes = rtf(300)
		const source = await openRtfSource(bytes)
		const reading = counting(() => undefined)
		const read = await openRtfSource(bytes, { signal: { aborted: false }, slicing: reading.slicing })

		if (source === undefined || read === undefined) throw new Error("not RTF")

		expect(reading.pauses()).toBeGreaterThan(5)
		expect(read.read.docx).toEqual(source.read.docx)

		const proving = counting(() => undefined)
		const plain = await proveRtf(source)
		const inSlices = await proveRtf(source, { signal: { aborted: false }, slicing: proving.slicing })

		expect(proving.pauses()).toBeGreaterThan(20)
		expect(inSlices).toEqual(plain)
	})

	it("ends a read when the document closes during its open: no further slice runs", async () => {
		const sliced = counting(count => {
			if (count === 5) server.close(1)
		})
		const server = docxRtfServer(sliced.slicing)

		expect(await server.open(1, rtf(1500))).toEqual({ ok: false, notRtf: false })
		expect(server.package(1)).toBeUndefined()

		for (let k = 0; k < 20; k++) await macrotask()

		expect(sliced.pauses()).toBe(5)
	})

	it("ends a proof when the document closes: no further slice runs", async () => {
		let closeAt = Infinity
		const sliced = counting(count => {
			if (count === closeAt) server.close(1)
		})
		const server = docxRtfServer(sliced.slicing)

		await server.open(1, rtf(1500))
		closeAt = sliced.pauses() + 5
		await expect(server.prove(1)).rejects.toBeInstanceOf(RtfAborted)

		for (let k = 0; k < 20; k++) await macrotask()

		expect(sliced.pauses()).toBe(closeAt)
	})

	it("ends a save when the document closes, and runs none queued after it", async () => {
		const bytes = rtf(1500)
		const editing = await editRtf(bytes)
		const target = paragraphs(editing.session.state.doc)[3]

		if (target === undefined) throw new Error("no paragraph")

		editing.session.dispatch(editing.session.state.tr.insertText("Edited ", target.pos + 1))

		const edits = await editsOf(editing)
		let closeAt = Infinity
		const sliced = counting(count => {
			if (count === closeAt) server.close(1)
		})
		const server = docxRtfServer(sliced.slicing)

		await server.open(1, bytes)
		closeAt = sliced.pauses() + 1

		const written = server.write(1, { parts: [...edits.parts], removed: edits.removed ?? [] })
		const queued = server.prove(1)

		await expect(written).rejects.toBeInstanceOf(RtfAborted)
		await expect(queued).rejects.toBeInstanceOf(RtfAborted)

		for (let k = 0; k < 20; k++) await macrotask()

		expect(sliced.pauses()).toBe(closeAt)
	})

	it("posts no answer for a proof its document's close ended", async () => {
		const posted: { t?: unknown; s?: unknown }[] = []
		let deliver: (data: unknown) => void = () => undefined

		serveDocxWorker(
			{
				post: message => posted.push(message as { t?: unknown; s?: unknown }),
				listen: onMessage => {
					deliver = onMessage
				}
			},
			() => {
				throw new Error("no engine here")
			}
		)

		const bytes = rtf(3000)

		deliver({ t: "rtf", s: 1, bytes: bytes.slice() })
		deliver({ t: "rtf", s: 2, bytes: bytes.slice() })

		for (let k = 0; k < 2000 && posted.filter(message => message.t === "rtfOpened").length < 2; k++) await macrotask()

		deliver({ t: "rtfProve", s: 1, id: 1 })
		deliver({ t: "rtfProve", s: 2, id: 1 })

		// Under way (Node hands a busy MessageChannel up to a thousand turns before a timer: closed from a microtask).
		for (let k = 0; k < 20; k++) await Promise.resolve()

		deliver({ t: "close", s: 1 })

		// The other document's proof, side by side, comes back; the closed one's never does.
		for (let k = 0; k < 5000 && !posted.some(message => message.t === "rtfProved" && message.s === 2); k++) await macrotask()

		for (let k = 0; k < 50; k++) await macrotask()

		expect(posted.some(message => message.t === "rtfProved" && message.s === 2)).toBe(true)
		expect(posted.some(message => message.t === "rtfProved" && message.s === 1)).toBe(false)
		deliver({ t: "close", s: 2 })
	})
})
