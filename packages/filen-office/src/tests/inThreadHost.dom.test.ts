// @vitest-environment happy-dom
import { TextSelection } from "prosemirror-state"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { prepareDocxSave } from "@filen/office/core/docx/write"
import { bytesSource } from "@filen/office/core/opc"
import { inThreadDocxHost, windowServices, type DocxEngineServices, type DocxEvent, type DocxPageData } from "@filen/office/dom/docx"
import { inThreadEditingHost } from "@filen/office/dom/docx/editing"
import { buildDocx, p } from "./docxSynthetic"
import { png } from "./security/generate"

// A picture whose header the viewer reads (one it cannot read is not served).
const IMAGE = png(4, 3)

// The engine as the mobile DOM component runs it: on the page's thread, with neither Worker nor
// MessageChannel nor OffscreenCanvas (an Expo DOM component on an older WebView), through windowServices.
// Every session entry must answer, and work must come in slices: the longest stretch between two yields is
// what the WebView's input waits for.

const PICTURE = "http://schemas.openxmlformats.org/drawingml/2006/picture"
const LOREM = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore. "

const DRAWING =
	'<w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="d"/><a:graphic><a:graphicData uri="' +
	PICTURE +
	'"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

// About 50 pages with the measurer below: long enough that open, layout and save each need many slices.
const LONG = buildDocx({
	body:
		'<w:p><w:bookmarkStart w:id="0" w:name="target"/><w:r><w:t>Start</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>' +
		DRAWING +
		Array.from({ length: 400 }, (_, k) => p(`${k}. ${LOREM.repeat(4)}`)).join(""),
	relationships: [{ id: "rIdImg", type: "image", target: "media/image1.png" }],
	parts: [{ name: "word/media/image1.png", data: IMAGE }]
})

// A 2D context whose glyphs are half an em wide (happy-dom has no canvas).
function context(): { font: string; fontKerning: string; measureText(text: string): { width: number } } {
	return {
		font: "16px serif",
		fontKerning: "normal",
		measureText(text: string) {
			return { width: text.length * Number(/([\d.]+)px/.exec(this.font)?.[1] ?? "16") * 0.5 }
		}
	}
}

class FontFaceStub {
	load(): Promise<this> {
		return Promise.resolve(this)
	}
}

// The longest run between yields per phase, and the yields counted.
class SliceClock {
	private last = 0
	longest = 0
	yields = 0

	start(): void {
		this.last = performance.now()
		this.longest = 0
		this.yields = 0
	}

	tick(): void {
		this.longest = Math.max(this.longest, performance.now() - this.last)
		this.yields++
	}

	resumed(): void {
		this.last = performance.now()
	}

	report(): string {
		return `${this.yields} yields, longest slice ${this.longest.toFixed(1)} ms`
	}
}

// windowServices as the DOM component gets them; a short `sliceMs` makes even a fast phase pause.
function services(clock: SliceClock, sliceMs?: number): DocxEngineServices {
	const base = windowServices(window, { read: () => Promise.resolve(new Uint8Array(4)) })

	return {
		...base,
		...(sliceMs === undefined ? {} : { sliceMs }),
		yield: async () => {
			clock.tick()
			await base.yield()
			clock.resumed()
		}
	}
}

class Pages {
	readonly pages: DocxPageData[] = []
	count = 0
	readonly done = Promise.withResolvers<void>()

	readonly listener = (event: DocxEvent): void => {
		if (event.type === "error") this.done.reject(new Error(event.message))
		else if (event.type === "pages") {
			for (const page of event.pages) this.pages[page.index] = page

			this.count = event.pageCount

			if (event.final) this.done.resolve()
		}
	}
}

describe("in-thread host without Worker, MessageChannel or OffscreenCanvas", () => {
	let timeouts = 0

	beforeEach(() => {
		vi.stubGlobal("Worker", undefined)
		vi.stubGlobal("MessageChannel", undefined)
		vi.stubGlobal("OffscreenCanvas", undefined)
		vi.stubGlobal("FontFace", FontFaceStub)
		Object.defineProperty(document, "fonts", { configurable: true, value: new Set() })
		vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockImplementation(() => context() as never)

		const setTimer = globalThis.setTimeout

		timeouts = 0
		vi.stubGlobal("setTimeout", (run: () => void, ms?: number) => {
			timeouts++

			return setTimer(run, ms)
		})
	})

	afterEach(() => {
		Reflect.deleteProperty(document, "fonts")
	})

	it("opens and lays out in slices, and answers every session entry", async ({ annotate }) => {
		const clock = new SliceClock()
		const pages = new Pages()
		const engineServices = services(clock)

		expect(engineServices.sliceMs).toBe(6)
		expect(engineServices.replayFonts).toBe(false)

		clock.start()

		const session = inThreadDocxHost(() => engineServices)(bytesSource(LONG), pages.listener)

		await pages.done.promise
		clock.tick()

		await annotate(`open + layout of ${pages.count} pages: ${clock.report()}`)
		expect(pages.count).toBeGreaterThan(30)
		expect(clock.yields).toBeGreaterThan(pages.count / 10)
		// Yields went through setTimeout (no MessageChannel), widths through a <canvas> (no OffscreenCanvas).
		expect(timeouts).toBeGreaterThanOrEqual(clock.yields - 1)
		expect(HTMLCanvasElement.prototype.getContext).toHaveBeenCalledWith("2d")

		const withObject = pages.pages.find(page => page.objectCount > 0)
		const [object] = await session.objects([withObject?.objectInts[0] ?? -1])

		expect(object?.kind).toBe("drawing")
		expect(Array.from((await session.media("word/media/image1.png"))?.bytes ?? [])).toEqual(Array.from(IMAGE))
		expect((await session.locate("target"))?.page).toBe(0)
		expect((await session.find("399.")).length).toBeGreaterThan(0)
		expect(await session.annotations?.()).not.toBeNull()

		session.close()
	})

	it("edits and saves in 1 ms slices; the save opens again with the edit", async ({ annotate }) => {
		const clock = new SliceClock()
		const pages = new Pages()
		const editing = inThreadEditingHost(() => services(clock, 1), {
			writer: { prepare: prepareDocxSave },
			author: "Mobile"
		})

		clock.start()

		const session = editing.host(bytesSource(LONG), pages.listener)
		const editor = await editing.editor

		await pages.done.promise
		clock.tick()
		await annotate(`editing open of ${pages.count} pages, 1 ms slices: ${clock.report()}`)

		const state = editor.state

		expect(editor.dispatch(state.tr.setSelection(TextSelection.create(state.doc, 2)).insertText("Mobile "))).toBe(true)
		expect(editing.setView("markup")).toBe(true)
		editing.setAuthor("Phone")

		const handle = editing.addMedia(new Uint8Array([137, 80, 78, 71, 9]), { contentType: "image/png" })

		expect(handle).toBeDefined()

		if (handle !== undefined) editing.releaseMedia(handle)

		await editor.whenSettled()

		const trial = await editing.prepare()

		expect(trial.changedParts).toContain("word/document.xml")

		clock.start()

		const written = await editing.save()

		clock.tick()
		await annotate(`save of ${written.bytes.length} bytes, 1 ms slices: ${clock.report()}`)
		expect(clock.yields).toBeGreaterThan(2)
		session.close()

		const reopened = new Pages()
		const viewer = inThreadDocxHost(() => services(new SliceClock()))(bytesSource(written.bytes), reopened.listener)

		await reopened.done.promise
		expect(reopened.count).toBe(pages.count)
		expect((await viewer.find("Mobile Start")).length).toBe(4)
		viewer.close()
	})
})
