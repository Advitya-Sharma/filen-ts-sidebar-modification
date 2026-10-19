// @vitest-environment happy-dom
import { describe, expect, it } from "vitest"
import type { DocxSessionApi } from "@filen/office/dom/docx"
import { DrawingList } from "../core/drawingml/displayList"
import { DocxMediaCache } from "../react/docx/media"
import { paintObjects } from "../react/docx/paint/objects"

// A picture's accessible name on the page (docxAltText.test.ts: where it comes from).

describe("picture alt text on the page", () => {
	it("names the picture's box on the page, as an image", async () => {
		const list = new DrawingList()

		list.width = 10
		list.height = 10

		const compact = list.compact()
		const drawing = {
			width: 10,
			height: 10,
			commands: compact.commands,
			coords: compact.coords,
			itemCount: compact.itemCount,
			itemFloats: compact.itemFloats,
			itemInts: compact.itemInts,
			paints: compact.paints,
			strokes: compact.strokes,
			images: compact.images,
			texts: compact.texts,
			shadows: compact.shadows
		}
		const session: DocxSessionApi = {
			objects: ids =>
				Promise.resolve(
					ids.map(id =>
						id === 1 ? { kind: "drawing" as const, id, drawing, alt: "A cat" } : { kind: "drawing" as const, id, drawing }
					)
				),
			media: () => Promise.resolve(null),
			locate: () => Promise.resolve(null),
			find: () => Promise.resolve(new Int32Array(0)),
			close: () => undefined
		}
		const media = new DocxMediaCache(session)
		const named = window.document.createElement("div")
		const plain = window.document.createElement("div")
		const release = paintObjects(
			[
				{ element: named, id: 1, kind: 0, width: 10, height: 10 },
				{ element: plain, id: 2, kind: 0, width: 10, height: 10 }
			],
			{ document: window.document, session, media, signal: { aborted: false } }
		)

		await new Promise(resolve => setTimeout(resolve, 0))

		expect(named.getAttribute("role")).toBe("img")
		expect(named.getAttribute("aria-label")).toBe("A cat")
		// The drawing itself stays out of the accessibility tree.
		expect(named.querySelector("svg")?.getAttribute("aria-hidden")).toBe("true")
		expect(plain.hasAttribute("role")).toBe(false)
		expect(plain.hasAttribute("aria-label")).toBe(false)

		release()
		media.dispose()
	})
})
