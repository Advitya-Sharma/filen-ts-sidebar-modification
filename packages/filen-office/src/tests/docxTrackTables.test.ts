import { existsSync, readFileSync, rmSync } from "node:fs"
import { join } from "node:path"
import type { Node } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { afterAll, describe, expect, it } from "vitest"
import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import { encodeUtf8, parseXml } from "@filen/office/core/xml"
import { DocxDocBuilder, DocxEditorSession, docxSchema, type DocxEditorCommand } from "@filen/office/core/docx/edit"
import { DocxVerifyError, saveDocx, type CheckInput } from "@filen/office/core/docx/write"
import { freshSlice } from "../core/docx/edit/invariants/core"
import { resolveChanges, storyReviewIndex, type ReviewScope } from "../core/docx/review/accept"
import { ReviewIndex, ReviewKind } from "../core/docx/review/index"
import { tableChecks } from "../core/docx/write/checks/tables"
import { CORPUS_DIR } from "./corpusDir"
import { openBytes } from "./docxEditFixtures"
import { buildDocx, p } from "./docxSynthetic"
import { SCHEMA_ARCHIVE, XMLLINT, schemaDirectory, validityErrors } from "./fuzz/harness"
import { editedLayout, pageDigest, savedDigest, type EditedLayout } from "./pagedFixtures"

// Tracked property changes of tables, rows and cells (w:tblPrChange, w:tblGridChange, w:trPrChange,
// w:tcPrChange), recorded as Word records them, and accepting and rejecting them and the source's own
// (w:tblPrExChange and w:cellMerge too).

const nodes = docxSchema.nodes
const decoder = new TextDecoder()
const AUTHOR = "Ann Author"

function tc(text: string, tcPr = ""): string {
	return `<w:tc>${tcPr === "" ? "" : `<w:tcPr>${tcPr}</w:tcPr>`}${p(text)}</w:tc>`
}

function tr(cells: readonly string[], trPr = ""): string {
	return `<w:tr>${trPr === "" ? "" : `<w:trPr>${trPr}</w:trPr>`}${cells.join("")}</w:tr>`
}

function tbl(grid: readonly number[], rows: readonly string[], tblPr = '<w:tblW w:w="0" w:type="auto"/>'): string {
	return `<w:tbl><w:tblPr>${tblPr}</w:tblPr><w:tblGrid>${grid.map(width => `<w:gridCol w:w="${width}"/>`).join("")}</w:tblGrid>${rows.join("")}</w:tbl>`
}

const W = (w: number): string => `<w:tcW w:w="${w}" w:type="dxa"/>`

// a1..c3, the first row 500 twips high.
const GRID3 = tbl(
	[2000, 2000, 2000],
	[1, 2, 3].map(row =>
		tr(
			["a", "b", "c"].map(column => tc(`${column}${row}`, W(2000))),
			row === 1 ? '<w:trHeight w:val="500"/>' : ""
		)
	)
)

const STYLES =
	'<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>' +
	'<w:style w:type="table" w:styleId="LightList"><w:name w:val="Light List"/></w:style>'

function docx(body: string): Uint8Array {
	return buildDocx({
		body: `${p("Before the table")}${body}${p("After the table")}`,
		parts: [
			{
				name: "word/styles.xml",
				contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
				data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${STYLES}</w:styles>`
			}
		],
		relationships: [{ id: "rIdStyles", type: "styles", target: "styles.xml" }]
	})
}

function textPos(doc: Node, text: string): number {
	let found = -1

	doc.descendants((node, pos) => {
		if (found >= 0) return false
		if (node.isText && node.text?.includes(text) === true) found = pos + node.text.indexOf(text)

		return true
	})

	if (found < 0) throw new Error(`No text ${text}`)

	return found
}

function select(session: DocxEditorSession, from: string, to: string = from): void {
	const doc = session.state.doc

	session.dispatch(session.state.tr.setSelection(TextSelection.create(doc, textPos(doc, from), textPos(doc, to) + (to === from ? 0 : 1))))
}

function track(session: DocxEditorSession): void {
	session.context.review.tracking = true
	session.context.identity.author = AUTHOR
}

function firstTable(doc: Node): Node {
	let table: Node | undefined

	doc.descendants(node => {
		if (table !== undefined) return false
		if (node.type === nodes.table) table = node

		return table === undefined
	})

	if (table === undefined) throw new Error("No table")

	return table
}

// Each row's cells as "text:span[:r|:c]".
function shape(table: Node): string[] {
	const out: string[] = []

	table.forEach(row => {
		const cells: string[] = []

		row.forEach(cell => {
			const vMerge = cell.attrs["vMerge"] as string | null

			cells.push(`${cell.textContent}:${String(cell.attrs["colspan"])}${vMerge === null ? "" : vMerge === "restart" ? ":r" : ":c"}`)
		})
		out.push(cells.join(" "))
	})

	return out
}

async function documentXml(bytes: Uint8Array): Promise<string> {
	return decoder.decode(await (await OpcPackage.open(bytesSource(bytes))).read("word/document.xml"))
}

async function reopened(bytes: Uint8Array): Promise<Node> {
	const opened = await openBytes(bytes)
	const body = opened.model.body

	if (body === undefined) throw new Error("No body")

	return firstTable(new DocxDocBuilder(opened.context).build(body).doc)
}

// Runs `steps` in a laid-out editor (pages checked against a fresh layout of the saved file after each
// resolution and at the end) and returns it with the saved bytes.
async function edited(
	bytes: Uint8Array,
	steps: readonly (DocxEditorCommand | ((session: DocxEditorSession) => void))[]
): Promise<{ bytes: Uint8Array; layout: EditedLayout }> {
	const layout = await editedLayout(bytes)

	for (const step of steps) {
		if (typeof step === "function") step(layout.session)
		else expect(await layout.edit(session => session.execute(step)), step.type).toBe(true)
	}

	const saved = await layout.save()

	expect(pageDigest(layout.layout)).toEqual(await savedDigest(saved))

	return { bytes: saved, layout }
}

// Resolves `scope` in the laid-out editor; returns what was left.
async function resolve(layout: EditedLayout, scope: ReviewScope, accept: boolean): Promise<number> {
	const outcome = resolveChanges(layout.session, storyReviewIndex(layout.session), scope, accept)

	if (outcome.tr !== null) expect(await layout.edit(outcome.tr)).toBe(true)

	return outcome.skipped
}

// The ids of the property changes the review index lists, in document order.
function propertyChanges(session: DocxEditorSession): string[] {
	const index = storyReviewIndex(session)
	const out: string[] = []

	for (let k = 0; k < index.size; k++) {
		const entry = index.entry(k)

		if (entry?.kind === ReviewKind.PROPERTIES && !out.includes(entry.id)) out.push(entry.id)
	}

	return out
}

const CHANGE = (local: string): RegExp => new RegExp(`<w:${local} w:id="\\d+" w:author="${AUTHOR}" w:date="[0-9T:-]+Z">`, "g")

describe("tracked table property changes", () => {
	it("records cell, row and table property changes with the properties before them, the first record staying", async () => {
		const { bytes, layout } = await edited(docx(GRID3), [
			track,
			session => select(session, "b2"),
			{ type: "table.cell.shading", fill: "FFEE00" },
			{ type: "table.cell.vAlign", align: "center" },
			session => select(session, "a1"),
			{ type: "table.row.header" },
			{ type: "table.align", align: "center" },
			{ type: "table.style", styleId: "LightList" }
		])
		const xml = await documentXml(bytes)

		// One w:tcPrChange (the first edit's record; the second adds to it) holding the cell's properties before.
		expect(xml.match(CHANGE("tcPrChange"))?.length).toBe(1)
		expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="FFEE00"/><w:vAlign w:val="center"/><w:tcPrChange')
		expect(xml).toMatch(/<w:tcPrChange [^>]*><w:tcPr><w:tcW w:w="2000" w:type="dxa"\/><\/w:tcPr><\/w:tcPrChange>/)
		expect(xml).toMatch(
			/<w:trPr><w:trHeight w:val="500"\/><w:tblHeader\/><w:trPrChange [^>]*><w:trPr><w:trHeight w:val="500"\/><\/w:trPr><\/w:trPrChange><\/w:trPr>/
		)
		// The table's: its style and alignment changed, one record of how it was.
		expect(xml.match(CHANGE("tblPrChange"))?.length).toBe(1)
		expect(xml).toMatch(
			/<w:tblPr><w:tblStyle w:val="LightList"\/><w:tblW w:w="0" w:type="auto"\/><w:jc w:val="center"\/><w:tblPrChange [^>]*><w:tblPr><w:tblW w:w="0" w:type="auto"\/><\/w:tblPr><\/w:tblPrChange><\/w:tblPr>/
		)
		// Three changes (one per command, the cell's two one record), reopened as source changes.
		expect(propertyChanges(layout.session)).toHaveLength(3)

		const back = await openBytes(bytes)

		expect(new DocxEditorSession(back.context).snapshot().review.revisionCount).toBe(3)
	})

	it("records column widths as w:tblGridChange and w:tcPrChange, and column insertions with the grid before", async () => {
		const widths = await edited(docx(GRID3), [
			track,
			session => select(session, "b2"),
			{ type: "table.column.width", index: 0, twips: 3000 }
		])
		const xml = await documentXml(widths.bytes)

		expect(xml).toMatch(
			/<w:tblGrid><w:gridCol w:w="3000"\/><w:gridCol w:w="1000"\/><w:gridCol w:w="2000"\/><w:tblGridChange w:id="\d+"><w:tblGrid><w:gridCol w:w="2000"\/><w:gridCol w:w="2000"\/><w:gridCol w:w="2000"\/><\/w:tblGrid><\/w:tblGridChange><\/w:tblGrid>/
		)
		// Six cells changed width: one record each, one change with the grid's.
		expect(xml.match(CHANGE("tcPrChange"))?.length).toBe(6)
		expect(propertyChanges(widths.layout.session)).toHaveLength(1)

		const columns = await edited(docx(GRID3), [track, session => select(session, "b2"), { type: "table.column.insertRight" }])
		const inserted = await documentXml(columns.bytes)

		expect(inserted.match(/<w:cellIns /g)?.length).toBe(3)
		expect(inserted).toContain("<w:tblGridChange")
		expect(inserted.match(CHANGE("tcPrChange"))?.length).toBe(9)
		// Rejected, the column goes and the widths come back: the source again.
		expect(await resolve(columns.layout, { all: true }, false)).toBe(0)
		expect(await documentXml(await columns.layout.save())).toBe(await documentXml(docx(GRID3)))
	})

	it("gives the source back byte for byte when every tracked property change is rejected", async () => {
		const steps: (DocxEditorCommand | ((session: DocxEditorSession) => void))[] = [
			track,
			session => select(session, "b2", "c3"),
			{ type: "table.cell.shading", fill: "00FF00" },
			{ type: "table.cell.vAlign", align: "bottom" },
			session => select(session, "a1"),
			{ type: "table.row.header" },
			{ type: "table.style", styleId: "TableGrid" },
			{ type: "table.align", align: "end" },
			{ type: "table.column.width", index: 1, twips: 2600 },
			{ type: "table.column.insertLeft" },
			{ type: "table.distributeColumns" }
		]
		const rejected = await edited(docx(GRID3), steps)

		expect(await resolve(rejected.layout, { all: true }, false)).toBe(0)
		expect(propertyChanges(rejected.layout.session)).toEqual([])

		const back = await rejected.layout.save()

		expect(pageDigest(rejected.layout.layout)).toEqual(await savedDigest(back))
		expect(await documentXml(back)).toBe(await documentXml(docx(GRID3)))

		// Accepted, the document is what the same edits make untracked.
		const accepted = await edited(docx(GRID3), steps)

		expect(await resolve(accepted.layout, { all: true }, true)).toBe(0)

		const untracked = await edited(
			docx(GRID3),
			steps.filter(step => step !== track)
		)

		expect(await documentXml(await accepted.layout.save())).toBe(await documentXml(untracked.bytes))
	})

	it("accepts and rejects one change by its id, the others staying", async () => {
		const { layout } = await edited(docx(GRID3), [
			track,
			session => select(session, "a1"),
			{ type: "table.cell.shading", fill: "112233" },
			session => select(session, "c3"),
			{ type: "table.cell.shading", fill: "445566" },
			{ type: "table.align", align: "center" }
		])
		const [first, second, third] = propertyChanges(layout.session)

		expect(third).toBeDefined()
		expect(await resolve(layout, { id: first as string }, true)).toBe(0)
		expect(propertyChanges(layout.session)).toEqual([second, third])
		expect(await resolve(layout, { id: second as string }, false)).toBe(0)
		expect(propertyChanges(layout.session)).toEqual([third])

		const xml = await documentXml(await layout.save())

		// a1 keeps its fill without a record, c3 has none, the table's change stays.
		expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="112233"/></w:tcPr>')
		expect(xml).not.toContain("445566")
		expect(xml.match(CHANGE("tblPrChange"))?.length).toBe(1)
		expect(xml).not.toContain("tcPrChange")

		// Undone, the rejection's record is back.
		expect(await layout.edit(session => session.execute({ type: "history.undo" }))).toBe(true)
		expect(propertyChanges(layout.session)).toHaveLength(2)
	})

	it("resolves the cells a selection takes in", async () => {
		const { layout } = await edited(docx(GRID3), [
			track,
			session => select(session, "a1", "c1"),
			{ type: "table.cell.shading", fill: "CCCCCC" }
		])
		const doc = layout.session.state.doc

		// a1 and b1 end inside the selection, c1 after it.
		expect(await resolve(layout, { from: textPos(doc, "a1"), to: textPos(doc, "c1") }, false)).toBe(0)

		const xml = await documentXml(await layout.save())

		expect(xml.match(/CCCCCC/g)?.length).toBe(1)
		expect(xml.match(/<w:tcPrChange /g)?.length).toBe(1)
	})

	it("keeps the review index current edit by edit", async () => {
		const layout = await editedLayout(docx(GRID3))
		const session = layout.session
		const index = storyReviewIndex(session)
		const steps: (DocxEditorCommand | ((session: DocxEditorSession) => void))[] = [
			track,
			session => select(session, "b2"),
			{ type: "table.cell.shading", fill: "FFEE00" },
			{ type: "table.align", align: "center" },
			{ type: "table.style", styleId: "LightList" },
			{ type: "table.row.header" },
			{ type: "table.column.width", index: 0, twips: 2500 },
			{ type: "table.column.insertRight" },
			{ type: "history.undo" },
			{ type: "table.distributeColumns" }
		]

		for (const step of steps) {
			if (typeof step === "function") step(session)
			else expect(await layout.edit(editor => editor.execute(step)), step.type).toBe(true)

			expect(index.entries()).toEqual(new ReviewIndex(session.state.doc).entries())
		}

		expect(await resolve(layout, { all: true }, false)).toBe(0)
		expect(index.entries()).toEqual([])
	})

	it("records nothing on rows and cells inserted as tracked changes, and needs an author", async () => {
		const { bytes } = await edited(docx(GRID3), [
			track,
			session => select(session, "b2"),
			{ type: "table.row.insertBelow" },
			session => {
				const doc = session.state.doc

				// The new row's middle cell.
				session.dispatch(session.state.tr.setSelection(TextSelection.near(doc.resolve(textPos(doc, "c2") + 8))))
			},
			{ type: "table.cell.shading", fill: "ABCDEF" }
		])
		const xml = await documentXml(bytes)

		expect(xml).toContain("ABCDEF")
		expect(xml).not.toContain("tcPrChange")

		const anonymous = await editedLayout(docx(GRID3))

		anonymous.context.review.tracking = true
		select(anonymous.session, "b2")
		expect(anonymous.session.execute({ type: "table.cell.shading", fill: "ABCDEF" })).toBe(false)
	})

	it("gives a pasted copy of a changed table no record of its own", async () => {
		const { layout } = await edited(docx(GRID3), [
			track,
			session => select(session, "a1"),
			{ type: "table.cell.shading", fill: "123456" },
			{ type: "table.align", align: "center" }
		])
		const session = layout.session
		const doc = session.state.doc
		let at = -1

		doc.forEach((node, offset) => {
			if (at < 0 && node.type === nodes.table) at = offset
		})

		const slice = doc.slice(at, at + firstTable(doc).nodeSize)

		session.context.review.tracking = false
		session.dispatch(session.state.tr.setSelection(TextSelection.create(doc, textPos(doc, "After"))))
		// As a paste or drop puts it in.
		expect(await layout.edit(session.state.tr.replaceSelection(freshSlice(slice)))).toBe(true)
		expect(propertyChanges(session)).toHaveLength(2)

		const bytes = await layout.save()
		const xml = await documentXml(bytes)

		expect(pageDigest(layout.layout)).toEqual(await savedDigest(bytes))
		expect(xml.match(/<w:tcPrChange /g)?.length).toBe(1)
		expect(xml.match(/<w:tblPrChange /g)?.length).toBe(1)
	})

	it("rejects a tracked change of a cell merged or split since: the width of the columns it spans then", async () => {
		const widths = (layout: EditedLayout): string[] =>
			[...firstTable(layout.session.state.doc).content.content].map(row =>
				[...row.content.content].map(cell => `${String(cell.attrs["colspan"])}/${String(cell.attrs["colwidth"])}`).join(" ")
			)
		// The second column 3000 wide (tracked: b1..b3 record 2000 and the grid before), then b2 merged with c2.
		const merged = await edited(docx(GRID3), [
			track,
			session => select(session, "b2"),
			{ type: "table.column.width", index: 1, twips: 3000 },
			session => select(session, "b2", "c2"),
			{ type: "table.cells.merge" }
		])

		// (null: the source's width.)
		expect(widths(merged.layout)).toEqual(["1/null 1/3000 1/1000", "1/null 2/4000", "1/null 1/3000 1/1000"])
		expect(await resolve(merged.layout, { all: true }, false)).toBe(0)
		// The grid as it was, the other cells their source widths; b2 the width of columns b and c in that grid.
		expect(widths(merged.layout)).toEqual(["1/null 1/null 1/null", "1/null 2/4000", "1/null 1/null 1/null"])

		const bytes = await merged.layout.save()

		expect(pageDigest(merged.layout.layout)).toEqual(await savedDigest(bytes))
		expect(await documentXml(bytes)).toMatch(/<w:tcW w:w="4000" w:type="dxa"\/><w:gridSpan w:val="2"\/>/)
		expect(await documentXml(bytes)).not.toMatch(/PrChange/)

		// b3's fill (tracked: its width 2000 recorded), then b3 split in two: rejected, the half keeps its own width.
		const split = await edited(docx(GRID3), [
			track,
			session => select(session, "b3"),
			{ type: "table.cell.shading", fill: "FFEE00" },
			session => select(session, "b3"),
			{ type: "table.cell.split", rows: 1, cols: 2 }
		])
		const before = widths(split.layout)[2]

		expect(before).toBe("1/null 1/1000 1/1000 1/null")
		expect(await resolve(split.layout, { all: true }, false)).toBe(0)
		expect(widths(split.layout)[2]).toBe(before)
		// b3 without the fill (the half split off keeps it, untracked, as a new cell copies its model's).
		expect(await documentXml(await split.layout.save())).toContain(
			'<w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">b3</w:t>'
		)
	})

	it("merges and splits untracked while tracking, as Word does", async () => {
		const { bytes, layout } = await edited(docx(GRID3), [
			track,
			session => select(session, "a1", "b2"),
			{ type: "table.cells.merge" },
			session => select(session, "c1"),
			{ type: "table.cell.split", rows: 1, cols: 2 }
		])
		const xml = await documentXml(bytes)

		expect(xml).not.toMatch(/cellMerge|cellIns|tcPrChange|tblGridChange/)
		expect(storyReviewIndex(layout.session).revisionCount).toBe(0)
	})
})

describe("the source's tracked table property changes", () => {
	// A table whose style, a row's height, a cell's fill and a cell's merge changed as tracked changes.
	const SOURCE = tbl(
		[3000, 3000],
		[
			tr(
				[
					tc(
						"One",
						`${W(3000)}<w:shd w:val="clear" w:color="auto" w:fill="FF0000"/><w:tcPrChange w:id="11" w:author="Bob"><w:tcPr>${W(3000)}</w:tcPr></w:tcPrChange>`
					),
					tc("Two", `${W(3000)}<w:vMerge w:val="restart"/>`)
				],
				'<w:trHeight w:val="900"/><w:trPrChange w:id="12" w:author="Bob"><w:trPr><w:trHeight w:val="400"/></w:trPr></w:trPrChange>'
			),
			tr([tc("Three", W(3000)), tc("Four", `${W(3000)}<w:vMerge/><w:cellMerge w:id="13" w:author="Bob" w:vMerge="cont"/>`)])
		],
		'<w:tblStyle w:val="LightList"/><w:tblW w:w="0" w:type="auto"/><w:tblPrChange w:id="10" w:author="Bob"><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr></w:tblPrChange>'
	)

	it("counts them and accepts them, the records going", async () => {
		const { layout } = await edited(docx(SOURCE), [])

		expect(propertyChanges(layout.session)).toHaveLength(4)
		expect(await resolve(layout, { all: true }, true)).toBe(0)
		expect(propertyChanges(layout.session)).toEqual([])
		expect(storyReviewIndex(layout.session).revisionCount).toBe(0)

		const bytes = await layout.save()

		expect(pageDigest(layout.layout)).toEqual(await savedDigest(bytes))

		const xml = await documentXml(bytes)

		expect(xml).not.toMatch(/PrChange|cellMerge/)
		expect(xml).toContain('<w:tblStyle w:val="LightList"/>')
		expect(xml).toContain('<w:trHeight w:val="900"/>')
		expect(xml).toContain('w:fill="FF0000"')
		expect(shape(await reopened(bytes))).toEqual(["One:1 Two:1:r", "Three:1 Four:1:c"])
	})

	it("rejects them, the properties before them back", async () => {
		const { layout } = await edited(docx(SOURCE), [])

		expect(await resolve(layout, { all: true }, false)).toBe(0)

		const bytes = await layout.save()

		expect(pageDigest(layout.layout)).toEqual(await savedDigest(bytes))

		const xml = await documentXml(bytes)

		expect(xml).not.toMatch(/PrChange|cellMerge/)
		expect(xml).toContain('<w:tblStyle w:val="TableGrid"/>')
		expect(xml).toContain('<w:trHeight w:val="400"/>')
		expect(xml).not.toContain("FF0000")
		// The cell merged into the one above as a tracked change is on its own again.
		expect(shape(await reopened(bytes))).toEqual(["One:1 Two:1:r", "Three:1 Four:1"])
		expect(shape(firstTable(layout.session.state.doc))).toEqual(["One:1 Two:1:r", "Three:1 Four:1"])
	})

	it("reads w:cellMerge as ECMA-376 defines it: w:vMerge applied, w:vMergeOrig removed (not merged when absent)", async () => {
		// The cell's merge as the file has it (the revised state) with its record; what accepting and rejecting leave.
		const cases: readonly [string, string, string, string][] = [
			// A merge recorded: rejected, the cell is on its own.
			['<w:vMerge/><w:cellMerge w:id="13" w:author="Bob" w:vMerge="cont"/>', "Four:1:c", "Four:1:c", "Four:1"],
			// A split recorded (the standard's example): rejected, the cell is merged again.
			['<w:cellMerge w:id="13" w:author="Bob" w:vMergeOrig="cont"/>', "Four:1", "Four:1", "Four:1:c"],
			// A split with no merge removed named: none was (the default is "rest"), both ways it stays split.
			['<w:cellMerge w:id="13" w:author="Bob" w:vMerge="rest"/>', "Four:1", "Four:1", "Four:1"],
			// Neither setting: nothing to apply or take back.
			['<w:vMerge/><w:cellMerge w:id="13" w:author="Bob"/>', "Four:1:c", "Four:1:c", "Four:1:c"]
		]

		for (const [tcPr, opened, accepted, rejected] of cases) {
			const table = tbl(
				[3000, 3000],
				[
					tr([tc("One", W(3000)), tc("Two", `${W(3000)}<w:vMerge w:val="restart"/>`)]),
					tr([tc("Three", W(3000)), tc("Four", `${W(3000)}${tcPr}`)])
				]
			)

			for (const accept of [true, false]) {
				const { layout } = await edited(docx(table), [])

				expect(shape(firstTable(layout.session.state.doc))[1], tcPr).toBe(`Three:1 ${opened}`)
				expect(await resolve(layout, { all: true }, accept)).toBe(0)

				const bytes = await layout.save()

				expect(pageDigest(layout.layout)).toEqual(await savedDigest(bytes))
				expect(await documentXml(bytes)).not.toContain("cellMerge")
				expect(shape(await reopened(bytes))[1], `${tcPr} ${String(accept)}`).toBe(`Three:1 ${accept ? accepted : rejected}`)
			}
		}
	})

	it("resolves one node's changes by the id the index gives them", async () => {
		const { layout } = await edited(docx(SOURCE), [])
		const ids = propertyChanges(layout.session)

		expect(await resolve(layout, { id: ids[1] as string }, false)).toBe(0)
		expect(propertyChanges(layout.session)).toEqual(ids.filter(id => id !== ids[1]))

		const xml = await documentXml(await layout.save())

		expect(xml.match(/PrChange w:id/g)?.length).toBe(2)
	})

	it("leaves a rejection that would leave a row wider than the grid as it is, counted", async () => {
		// The cell spanned both columns before its change; its row has another cell now.
		const wide = tbl(
			[3000, 3000],
			[
				tr([
					tc(
						"One",
						`${W(3000)}<w:tcPrChange w:id="21" w:author="Bob"><w:tcPr>${W(6000)}<w:gridSpan w:val="2"/></w:tcPr></w:tcPrChange>`
					),
					tc("Two", W(3000))
				])
			]
		)
		const { layout } = await edited(docx(wide), [])

		expect(await resolve(layout, { all: true }, false)).toBe(1)
		expect(propertyChanges(layout.session)).toHaveLength(1)
		expect(shape(firstTable(layout.session.state.doc))).toEqual(["One:1 Two:1"])
		// Accepted, it goes.
		expect(await resolve(layout, { all: true }, true)).toBe(0)
		expect(await documentXml(await layout.save())).not.toContain("tcPrChange")
	})

	it("rejects a column insertion one of whose cells a merge took in not at all, rather than half", async () => {
		const steps = [
			track,
			(session: DocxEditorSession) => select(session, "b2"),
			{ type: "table.column.insertRight" } as const,
			(session: DocxEditorSession) => {
				const doc = session.state.doc

				session.dispatch(session.state.tr.setSelection(TextSelection.create(doc, textPos(doc, "b2"), textPos(doc, "c2") + 1)))
			},
			{ type: "table.cells.merge" } as const
		]

		for (const reopen of [false, true]) {
			const { bytes, layout } = await edited(docx(GRID3), steps)
			const target = reopen ? await editedLayout(bytes) : layout
			const before = shape(firstTable(target.session.state.doc))
			const pending = storyReviewIndex(target.session).size

			const xml = await documentXml(await target.save())

			expect(await resolve(target, { all: true }, false)).toBeGreaterThan(0)
			expect(await documentXml(await target.save())).toBe(xml)
			expect(shape(firstTable(target.session.state.doc))).toEqual(before)
			expect(storyReviewIndex(target.session).size).toBe(pending)

			const saved = await target.save()

			expect(pageDigest(target.layout)).toEqual(await savedDigest(saved))
		}
	})

	it.skipIf(!existsSync(join(CORPUS_DIR, "docx", "oxpt__RA001-Tracked-Revisions-01.docx")))(
		"resolves Word's own table property changes in RA001",
		async () => {
			for (const accept of [true, false]) {
				const bytes = new Uint8Array(readFileSync(join(CORPUS_DIR, "docx", "oxpt__RA001-Tracked-Revisions-01.docx")))
				const opened = await openBytes(bytes)
				const session = new DocxEditorSession(opened.context)
				const outcome = resolveChanges(session, storyReviewIndex(session), { all: true }, accept)

				expect(session.dispatch(outcome.tr as never)).toBe(true)
				// Left: only inserted paragraph marks ending cells (no paragraph after them to join).
				expect(propertyChanges(session)).toEqual([])
				expect(outcome.skipped).toBe(30)

				const result = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [session] })
				const xml = await documentXml(result.bytes as Uint8Array)

				expect(xml).not.toMatch(/PrChange|tblGridChange|cellMerge|<w:trPr>(?:(?!<\/w:trPr>).)*<w:(?:ins|del) /)
				// Rejected, the row is as high as it was before its change.
				if (!accept) expect(xml).toContain('<w:trHeight w:val="125"/>')
			}
		}
	)
})

describe("gate 2 checks of tracked table property changes", () => {
	async function check(written: string): Promise<unknown> {
		const opened = await openBytes(docx(GRID3))
		const part = opened.model.parts[opened.model.body?.part ?? 0]
		const document = parseXml(encodeUtf8(await documentXml(docx(written))))
		const input = { original: opened.pkg, changes: { stories: [] } } as unknown as CheckInput

		try {
			await tableChecks[0]?.part?.(input, { name: "word/document.xml", part, document, skeleton: false })

			return undefined
		} catch (error) {
			return error instanceof DocxVerifyError ? error.message : error
		}
	}

	const fault = "word/document.xml: a tracked table, row or cell change has no id, author or previous properties"

	it("refuses records without an id, an author or the properties before them", async () => {
		const cell = (tcPr: string): string => tbl([2000], [tr([tc("a", tcPr)])])

		expect(await check(cell(`${W(2000)}<w:tcPrChange w:id="5" w:author="Ann"><w:tcPr/></w:tcPrChange>`))).toBeUndefined()
		expect(await check(cell(`${W(2000)}<w:tcPrChange w:id="5"><w:tcPr/></w:tcPrChange>`))).toBe(fault)
		expect(await check(cell(`${W(2000)}<w:tcPrChange w:id="5" w:author="Ann"/>`))).toBe(fault)
		expect(await check(cell('<w:cellMerge w:id="5" w:vMerge="cont"/>'))).toBe(fault)
		expect(
			await check(
				'<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:tblGridChange w:id="4"><w:tblGrid/></w:tblGridChange></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>'
			)
		).toBeUndefined()
		expect(
			await check(
				'<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:tblGridChange><w:tblGrid/></w:tblGridChange></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>'
			)
		).toBe(fault)
		expect(await check(tbl([2000], [tr([tc("a")], '<w:trPrChange w:id="3" w:author=""><w:trPr/></w:trPrChange>')]))).toBe(fault)
		expect(await check(tbl([2000], [tr([tc("a")])], '<w:tblPrChange w:id="3" w:author="Ann"/>'))).toBe(fault)
	})
})

describe.skipIf(!XMLLINT || !existsSync(SCHEMA_ARCHIVE))("tracked table property changes against the ECMA-376 schemas", () => {
	const directory = XMLLINT && existsSync(SCHEMA_ARCHIVE) ? schemaDirectory() : ""

	afterAll(() => {
		if (directory !== "") rmSync(directory, { recursive: true, force: true })
	})

	it("writes valid WordprocessingML", { timeout: 120_000 }, async () => {
		const { bytes } = await edited(docx(GRID3), [
			track,
			session => select(session, "b2", "c3"),
			{ type: "table.cell.shading", fill: "00FF00" },
			{ type: "table.column.width", index: 1, twips: 2600 },
			session => select(session, "a1"),
			{ type: "table.row.header" },
			{ type: "table.style", styleId: "TableGrid" },
			{ type: "table.column.insertRight" }
		])
		const written = await (await OpcPackage.open(bytesSource(bytes))).read("word/document.xml")

		expect([...validityErrors(join(directory, "wml.xsd"), written)]).toEqual([])
		expect((await reopened(bytes)).attrs["grid"]).toHaveLength(4)
	})
})
