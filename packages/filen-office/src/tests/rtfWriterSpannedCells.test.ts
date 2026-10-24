import { strFromU8, unzipSync } from "fflate"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { readRtf } from "../core/rtf/index"
import { editRtf, editsOf, latin1, paragraphs, saveRtf } from "./rtfWriteFixtures"

// A cell spanning grid columns is written as one wider cell (its \cellx at the span's right edge), as Word and
// LibreOffice write it, when the table's other rows have the edges inside the span; a reader rebuilds the span from
// those rows. The table then reads back editable, with the cells it had.

const HEADER = "{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\froman Times New Roman;}}"

const SPANNED = latin1(
	HEADER +
		"\\pard\\plain\\f0\\fs24 Before the table\\par\n" +
		"\\trowd\\trgaph108\\trleft0\\cellx6000\\pard\\plain\\intbl Title across\\cell\\row\n" +
		"\\trowd\\trgaph108\\trleft0\\cellx3000\\cellx6000\\pard\\plain\\intbl Left cell\\cell Right cell\\cell\\row\n" +
		"\\pard\\plain\\f0\\fs24 After the table\\par\n}"
)

async function bodyXml(bytes: Uint8Array): Promise<string> {
	const read = await readRtf(bytes)

	if (!read.ok) throw new Error(read.reason)

	return strFromU8(unzipSync(read.docx)["word/document.xml"] ?? new Uint8Array())
}

// Each row's cells as their grid spans, and the table's grid.
function shape(xml: string): { grid: string[]; rows: number[][] } {
	const table = /<w:tbl>.*<\/w:tbl>/su.exec(xml)?.[0] ?? ""
	const grid = [...table.matchAll(/<w:gridCol w:w="(\d+)"\/>/gu)].map(match => match[1] ?? "")
	const rows = [...table.matchAll(/<w:tr[ >].*?<\/w:tr>/gsu)].map(row =>
		[...row[0].matchAll(/<w:tc>.*?<\/w:tc>/gsu)].map(cell => Number(/<w:gridSpan w:val="(\d+)"\/>/u.exec(cell[0])?.[1] ?? "1"))
	)

	return { grid, rows }
}

async function insertRowBelow(bytes: Uint8Array, cell: number): Promise<{ written: Uint8Array; edited: string }> {
	const editing = await editRtf(bytes)

	expect(editing.locked).toBe(0)

	const at = paragraphs(editing.session.state.doc, true)[cell]

	if (at === undefined) throw new Error("no cell")

	const session = editing.session

	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, at.pos + 1)))
	expect(await editing.editor.execute({ type: "table.row.insertBelow" })).toBe(true)

	const edited = strFromU8((await editsOf(editing)).parts.get("word/document.xml") ?? new Uint8Array())

	return { written: (await saveRtf(editing)).bytes, edited }
}

describe("RTF cells spanning grid columns", () => {
	it("writes a new row's spanning cell as one wider cell, so the table reopens editable with the same cells", async () => {
		const { written, edited } = await insertRowBelow(SPANNED, 0)
		const out = strFromU8(written, true)

		expect(out).not.toContain("\\clmgf")
		expect(out).not.toContain("\\clmrg")
		expect(shape(edited).rows).toEqual([[2], [2], [1, 1]])
		expect(shape(await bodyXml(written))).toEqual(shape(edited))
		expect((await editRtf(written)).locked).toBe(0)
	})

	it("writes a span no other row divides as merged cells, so the grid keeps its columns", async () => {
		const editing = await editRtf(
			latin1(
				HEADER +
					"\\trowd\\trgaph108\\trleft0\\cellx3000\\cellx6000\\pard\\plain\\intbl Left\\cell Right\\cell\\row\n" +
					"\\pard\\plain\\f0\\fs24 After the table\\par\n}"
			)
		)
		const session = editing.session
		const cells = paragraphs(session.state.doc, true)
		const [left, right] = cells

		if (left === undefined || right === undefined) throw new Error("no cells")

		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, left.pos + 1, right.pos + 2)))
		expect(await editing.editor.execute({ type: "table.cells.merge" })).toBe(true)

		const edited = strFromU8((await editsOf(editing)).parts.get("word/document.xml") ?? new Uint8Array())
		const written = (await saveRtf(editing)).bytes

		expect(shape(edited)).toEqual({ grid: ["3000", "3000"], rows: [[2]] })
		expect(strFromU8(written, true)).toContain("\\clmgf")
		expect(shape(await bodyXml(written))).toEqual(shape(edited))
	})
})
