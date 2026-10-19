import { describe, expect, it } from "vitest"
import { docxSchema, type DocxEditorSession } from "@filen/office/core/docx/edit"
import { TextSelection } from "prosemirror-state"
import { laidTable, tableLook } from "../core/docx/reflow/tables"
import { Item, RevisionView } from "../core/docx/model/kinds"
import { buildDocx } from "./docxSynthetic"
import { loadDocx } from "./layoutFixtures"
import { editedLayout } from "./pagedFixtures"

// Mobile view draws a table's borders and shading from the table the pages lay out: with rows an edit wrote
// again alone (LayoutEdits.rows), as the saved file draws them.

const cell = (text: string, tcPr = ""): string =>
	`<w:tc>${tcPr === "" ? "" : `<w:tcPr>${tcPr}</w:tcPr>`}<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`

function bandedTable(): Uint8Array {
	const rows = Array.from({ length: 40 }, (_, r) =>
		r === 0
			? `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell("Head a")}${cell("Head b")}</w:tr>`
			: `<w:tr>${cell(`row ${r}`, r % 7 === 3 ? '<w:shd w:val="clear" w:color="auto" w:fill="FFCC00"/>' : "")}${cell(`b ${r}`)}</w:tr>`
	).join("")
	const style =
		'<w:style w:type="table" w:styleId="Banded"><w:name w:val="Banded"/><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblBorders><w:top w:val="single" w:sz="8"/><w:bottom w:val="single" w:sz="8"/><w:insideH w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>' +
		'<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="DDEEFF"/></w:tcPr></w:tblStylePr><w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="333333"/></w:tcPr></w:tblStylePr></w:style>'

	return buildDocx({
		body:
			`<w:tbl><w:tblPr><w:tblStyle w:val="Banded"/><w:tblW w:w="5000" w:type="pct"/><w:tblLook w:firstRow="1" w:noHBand="0"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>${rows}</w:tbl>` +
			'<w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>',
		parts: [
			{
				name: "word/styles.xml",
				contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
				data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${style}</w:styles>`
			}
		],
		relationships: [{ id: "rIdS", type: "styles", target: "styles.xml" }]
	})
}

function at(session: DocxEditorSession, text: string): number {
	let found = -1

	session.state.doc.descendants((node, pos) => {
		if (found < 0 && node.type === docxSchema.nodes.paragraph && node.textContent.startsWith(text)) found = pos + 1 + text.length

		return found < 0
	})

	if (found < 0) throw new Error(`no ${text}`)

	return found
}

describe("Mobile view's table look with rows an edit wrote again", () => {
	it("draws each row as the saved file does after rows are inserted, deleted and split", async () => {
		const edited = await editedLayout(bandedTable())
		const session = edited.session
		const select = (pos: number): void => {
			session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, pos)))
		}

		select(at(session, "row 5"))
		expect(await edited.edit(s => s.execute({ type: "table.row.insertBelow" }))).toBe(true)
		select(at(session, "row 20"))
		expect(await edited.edit(s => s.execute({ type: "table.row.delete" }))).toBe(true)
		await edited.edit(session.state.tr.split(at(session, "row 30")))

		// The table as laid out, with its rows (the shown look's), against the saved file's table.
		const entry = edited.story.entries.find(candidate => candidate.node.type === docxSchema.nodes.table)
		const laid = laidTable(edited.edits.model, entry?.laid[0], edited.edits.rows)

		if (laid === undefined) throw new Error("no table")

		expect(laid.edited).toBeDefined()

		const loaded = await loadDocx(await edited.save())
		const body = loaded.model.parts[loaded.model.body?.part ?? -1]
		let item = -1

		for (let k = 0; body !== undefined && k < body.kinds.length && item < 0; k++) if (body.kinds[k] === Item.Table) item = k

		for (const view of [RevisionView.Markup, RevisionView.Final]) {
			const shown = tableLook(laid.part, laid.item, 2, edited.context.resolver, edited.context.store, view, laid.edited)

			if (body === undefined) throw new Error("no body")

			const saved = tableLook(body, item, 2, loaded.resolver, loaded.store, view)

			expect(shown.cells).toEqual(saved.cells)
		}
	})
})
