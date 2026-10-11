import type { Node } from "prosemirror-model"
import { describe, expect, it } from "vitest"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import { LineFloat } from "@filen/office/core/text"
import { NUMBERING_PART } from "./docxEditFixtures"
import { buildDocx, p } from "./docxSynthetic"
import { pagedEditor as measuredEditor } from "./pagedFixtures"
import { pagedEditor } from "./pagedDomFixtures"

// Caret motion over the pages: by lines through columns and table cells, on screen through right-to-left
// text, at line wraps; list markers as text and the positions a page shows.

const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const para = (text: string, pPr = ""): string => `<w:p>${pPr === "" ? "" : `<w:pPr>${pPr}</w:pPr>`}${run(text)}</w:p>`
const cell = (inner: string): string => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`
const table = (rows: readonly (readonly string[])[]): string =>
	'<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
	rows.map(row => `<w:tr>${row.map(cell).join("")}</w:tr>`).join("") +
	"</w:tbl>"

// The position inside the textblock whose text starts with `text`.
function at(doc: Node, text: string, offset = 0): number {
	let found = -1

	doc.descendants((node, pos) => {
		if (found >= 0) return false
		if (node.isTextblock && node.textContent.startsWith(text)) found = pos + 1 + offset

		return found < 0
	})

	if (found < 0) throw new Error(`no paragraph ${text}`)

	return found
}

function textAt(editor: PagedEditor, pos: number): string {
	return editor.state.doc.resolve(pos).parent.textContent
}

describe("paged motion", () => {
	it("moves down from a short centred line to a short line below it", async () => {
		const { editor } = await pagedEditor(buildDocx({ body: para("Top") + para("Hi", '<w:jc w:val="center"/>') + para("ok") }))
		const doc = editor.state.doc
		const from = at(doc, "Hi", 1)
		const caret = editor.geometry.locate(from)
		const next = editor.geometry.verticalMove(from, 1, caret?.x ?? 0, 1)

		expect(next).toBeDefined()
		expect(textAt(editor, next?.pos ?? 0)).toBe("ok")

		const back = editor.geometry.verticalMove(next?.pos ?? 0, 1, caret?.x ?? 0, -1)

		expect(textAt(editor, back?.pos ?? 0)).toBe("Hi")
	})

	it("moves by lines within a table cell's column, not into the cell beside it", async () => {
		const tall = "word ".repeat(40)
		const { editor } = await pagedEditor(
			buildDocx({
				body:
					table([
						[para("Left one"), para(tall)],
						[para("Left two"), para("Right two")]
					]) + para("After")
			})
		)
		const doc = editor.state.doc
		const geometry = editor.geometry
		const from = at(doc, "Left one", 2)
		const caret = geometry.locate(from)
		const page = editor.page(caret?.page ?? 0)

		if (caret === undefined || page === undefined) throw new Error("no caret")

		// The column is the left cell's: narrower than the page's text and clear of the right cell.
		const column = geometry.columnBox(page, caret.line)
		const right = geometry.locate(at(doc, "word"))

		expect(column.right).toBeLessThan(right?.x ?? 0)

		const down = geometry.verticalMove(from, 1, caret.x, 1)

		expect(textAt(editor, down?.pos ?? 0)).toBe("Left two")

		const below = geometry.verticalMove(down?.pos ?? 0, 1, caret.x, 1)

		expect(textAt(editor, below?.pos ?? 0)).toBe("After")

		// From the tall cell's last line down: the next row's right cell.
		const tallEnd = at(doc, "word", tall.length - 2)
		const end = geometry.locate(tallEnd)
		const fromTall = geometry.verticalMove(tallEnd, 1, end?.x ?? 0, 1)

		expect(textAt(editor, fromTall?.pos ?? 0)).toBe("Right two")
	})

	it("moves down a text column, then to the top of the next", async () => {
		const body =
			Array.from({ length: 40 }, (_, k) => para(`Line ${String(k)} ${"text ".repeat(12)}`)).join("") +
			'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:cols w:num="2" w:space="720"/></w:sectPr>'
		const { editor } = await pagedEditor(buildDocx({ body }))
		const geometry = editor.geometry
		const page = editor.page(0)

		if (page === undefined) throw new Error("no page")

		const middle = page.width / 2
		let hit = geometry.hitTest(0, 120, 100)
		let previousTop = -1
		let switched = false

		for (let moves = 0; hit !== undefined && hit.page === 0 && moves < 400; moves++) {
			const box = geometry.locate(hit.pos, hit.affinity)

			if (box === undefined) break

			if (box.x > middle) {
				// The first line of the second column, at the top.
				expect(box.top).toBeLessThan(previousTop)
				switched = true

				break
			}

			expect(box.top).toBeGreaterThan(previousTop)
			previousTop = box.top
			hit = geometry.verticalMove(hit.pos, hit.affinity, 120, 1)
		}

		expect(switched).toBe(true)
	})

	it("steps on screen through right-to-left text", async () => {
		const { editor } = await pagedEditor(buildDocx({ body: para("שלום עולם יפה", "<w:bidi/>") + para("plain left") }))
		const doc = editor.state.doc
		const geometry = editor.geometry
		const from = at(doc, "שלום", 4)
		const caret = geometry.locate(from)

		expect(caret?.rtl).toBe(true)

		const right = geometry.visualMove(from, 1, 1)
		const left = geometry.visualMove(from, 1, -1)

		// Right on screen is back in the text, left is forward.
		expect(right?.pos).toBe(from - 1)
		expect(left?.pos).toBe(from + 1)
		expect(geometry.locate(right?.pos ?? 0)?.x ?? 0).toBeGreaterThan(caret?.x ?? 0)
		expect(geometry.locate(left?.pos ?? 0)?.x ?? 0).toBeLessThan(caret?.x ?? 0)

		// Left to right lines are left to the browser's own motion.
		expect(geometry.visualMove(at(doc, "plain", 2), 1, 1)).toBeUndefined()
	})

	it("tells where lines wrap", async () => {
		const { editor } = await pagedEditor(buildDocx({ body: para("word ".repeat(80)) }))
		const geometry = editor.geometry
		const page = editor.page(0)

		if (page === undefined) throw new Error("no page")

		const end = geometry.atLine(page, 0, 10_000)

		expect(geometry.wrapsAt(end?.pos ?? 0)).toBe(true)
		expect(geometry.wrapsAt((end?.pos ?? 0) - 2)).toBe(false)
	})

	it("reports the positions each page shows", async () => {
		const { editor } = await pagedEditor(
			buildDocx({ body: Array.from({ length: 60 }, (_, k) => para(`Paragraph ${String(k)} ${"lorem ipsum ".repeat(20)}`)).join("") })
		)
		const geometry = editor.geometry
		const first = geometry.pageRange(0)
		const second = geometry.pageRange(1)

		expect(first?.from).toBe(0)
		expect(second).toBeDefined()
		expect(second?.from ?? 0).toBeGreaterThan(first?.from ?? 0)
		// A paragraph split over the pages is on both.
		expect(second?.from ?? 0).toBeLessThanOrEqual(first?.to ?? 0)
		expect(geometry.pageRange(editor.pageCount + 5)).toBeUndefined()

		// Lines of the second page lie within its range.
		const page = editor.page(1)
		const top = page?.text.lineFloats[LineFloat.TOP] ?? 0
		const hit = geometry.hitTest(1, 200, top + 1)

		expect(hit?.pos ?? -1).toBeGreaterThanOrEqual(second?.from ?? 0)
		expect(hit?.pos ?? Infinity).toBeLessThanOrEqual(second?.to ?? 0)
	})

	it("gives list markers as text", async () => {
		const numbering = NUMBERING_PART(
			'<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
				'<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
		)
		const item = (text: string): string => para(text, '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>')
		const { editor } = await measuredEditor(
			buildDocx({
				body: item("First item") + item("Second item") + p("Plain"),
				parts: [numbering],
				relationships: [{ id: "rIdN", type: "numbering", target: "numbering.xml" }]
			})
		)
		const doc = editor.state.doc
		const geometry = editor.geometry

		expect(geometry.markerText(at(doc, "First item"))).toMatch(/^1\.\s*$/)
		expect(geometry.markerText(at(doc, "Second item", 3))).toMatch(/^2\.\s*$/)
		expect(geometry.markerText(at(doc, "Plain"))).toBe("")
	})
})
