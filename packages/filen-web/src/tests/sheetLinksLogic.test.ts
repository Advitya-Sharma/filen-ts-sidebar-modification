import { describe, expect, it } from "vitest"
import { CellFlag, type SheetSummary } from "@filen/office/core/sheet"
import {
	annotationOf,
	cardPlace,
	followsLink,
	linkAction,
	linkTarget,
	noteStart,
	placeName,
	type LinkForm,
	type SheetLink
} from "@/features/spreadsheet/comments/annotations.logic"
import { refusalKey } from "@/features/spreadsheet/lib/sheetMessages.logic"
import { preview } from "@/locales/en/preview"

const SHEETS: SheetSummary[] = [
	{ id: 1, name: "One", state: "visible", kind: "worksheet", ready: true },
	{ id: 2, name: "My Sheet", state: "visible", kind: "worksheet", ready: true },
	{ id: 3, name: "A1", state: "visible", kind: "worksheet", ready: true }
]

const range = (startRow: number, startCol: number, endRow = startRow, endCol = startCol) => ({ startRow, startCol, endRow, endCol })

function link(fields: Partial<SheetLink>): SheetLink {
	return { url: null, location: null, place: null, tooltip: null, display: null, range: null, formula: false, ...fields }
}

const WEB: LinkForm = { kind: "web", address: "", sheet: "One", reference: "A1", text: "", tooltip: "" }

describe("sheet annotations", () => {
	it("says what a cell has, a comment before its placeholder note", () => {
		expect(annotationOf(CellFlag.NOTE | CellFlag.THREADED_COMMENT)).toBe("thread")
		expect(annotationOf(CellFlag.NOTE | CellFlag.HYPERLINK)).toBe("note")
		expect(annotationOf(CellFlag.HYPERLINK | CellFlag.NUMBER)).toBe("link")
		expect(annotationOf(CellFlag.NUMBER)).toBeNull()
	})

	it("starts a new note with the author's line, as Excel does", () => {
		expect(noteStart("Jan")).toBe("Jan:\n")
		expect(noteStart(null)).toBe("")
	})

	it("follows a link with Command on a Mac and Ctrl elsewhere", () => {
		expect(followsLink({ ctrlKey: false, metaKey: true }, true)).toBe(true)
		expect(followsLink({ ctrlKey: true, metaKey: false }, true)).toBe(false)
		expect(followsLink({ ctrlKey: true, metaKey: false }, false)).toBe(true)
	})

	it("opens only what the shared link policy allows, and goes to places in the workbook", () => {
		expect(linkAction(link({ url: " https://filen.io/x " }))).toEqual({ kind: "open", href: "https://filen.io/x" })
		expect(linkAction(link({ url: "mailto:a@b.c" }))).toEqual({ kind: "open", href: "mailto:a@b.c" })

		for (const url of [
			"javascript:alert(1)",
			"java\tscript:x",
			"file:///etc/passwd",
			"data:text/html,x",
			"https:/nohost",
			"#local",
			"/relative"
		])
			expect(linkAction(link({ url })), url).toEqual({ kind: "blocked" })

		expect(linkAction(link({ location: "One!B2", place: { sheet: 1, range: range(1, 1) } }))).toEqual({
			kind: "go",
			place: { sheet: 1, range: range(1, 1) }
		})
		expect(linkAction(link({ location: "Gone!A1" }))).toEqual({ kind: "nowhere" })
		expect(linkAction(link({ formula: true }))).toEqual({ kind: "unknown" })
		expect(linkAction(null)).toEqual({ kind: "unknown" })
	})

	it("names a place as Excel does, quoting sheet names that need it", () => {
		expect(placeName(SHEETS, { sheet: 1, range: range(1, 1) })).toBe("One!B2")
		expect(placeName(SHEETS, { sheet: 2, range: range(0, 0, 4, 27) })).toBe("'My Sheet'!A1:AB5")
		expect(placeName(SHEETS, { sheet: 3, range: range(0, 0) })).toBe("'A1'!A1")
	})

	it("reads the link dialog: schemes added as Excel adds them, the policy deciding, places checked", () => {
		expect(linkTarget({ ...WEB, address: "filen.io/docs" })).toEqual({ ok: true, target: { url: "https://filen.io/docs" } })
		expect(linkTarget({ ...WEB, address: "jan@filen.io" })).toEqual({ ok: true, target: { url: "mailto:jan@filen.io" } })
		expect(linkTarget({ ...WEB, address: "tel:+4930123" })).toEqual({ ok: true, target: { url: "tel:+4930123" } })
		expect(linkTarget({ ...WEB, address: "javascript:alert(1)" })).toEqual({ ok: false, field: "address" })
		expect(linkTarget({ ...WEB, address: "" })).toEqual({ ok: false, field: "address" })
		expect(linkTarget({ ...WEB, kind: "place", sheet: "My Sheet", reference: "$B$2" })).toEqual({
			ok: true,
			target: { location: "'My Sheet'!$B$2" }
		})
		expect(linkTarget({ ...WEB, kind: "place", sheet: "One", reference: "Totals" })).toEqual({
			ok: true,
			target: { location: "Totals" }
		})
		expect(linkTarget({ ...WEB, kind: "place", reference: "A1:" })).toEqual({ ok: false, field: "reference" })
	})

	it("puts a card past the cell's trailing edge, on the other side when there is no room, inside the grid", () => {
		const bounds = { width: 800, height: 600 }

		expect(cardPlace({ x: 100, y: 50, width: 64, height: 20 }, 288, 120, bounds, false)).toEqual({ left: 170, top: 50 })
		expect(cardPlace({ x: 700, y: 50, width: 64, height: 20 }, 288, 120, bounds, false)).toEqual({ left: 406, top: 50 })
		expect(cardPlace({ x: 700, y: 50, width: 64, height: 20 }, 288, 120, bounds, true)).toEqual({ left: 406, top: 50 })
		expect(cardPlace({ x: 100, y: 590, width: 64, height: 20 }, 288, 120, bounds, false)).toEqual({ left: 170, top: 474 })
	})

	it("has a message for every refusal the annotations give", () => {
		for (const reason of ["notesUnavailable", "authorNeeded", "notOwnComment", "linkTarget"] as const)
			expect(preview).toHaveProperty(refusalKey({ reason }, "xlsx"))
	})
})
