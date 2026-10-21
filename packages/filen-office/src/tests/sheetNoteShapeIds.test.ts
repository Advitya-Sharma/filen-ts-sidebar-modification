import { describe, expect, it } from "vitest"
import { savedBytes } from "./sheetSessionFixtures"
import { NOTES, VML_HEAD, changeOf, entries, noteShape, openAnnotated, settle } from "./sheetNotesFixtures"

// Excel's markup for note shapes in workbooks with threaded comments: the id in o:spid, id itself "_x005f_x0000_s-1".
function excelShape(n: number, spid: number, row: number, col: number): string {
	return noteShape(spid, row, col).replace(
		`id="_x0000_s${String(spid)}"`,
		`id="_x005f_x0000_s${String(n)}" o:spid="_x0000_s${String(spid)}"`
	)
}

const VML = NOTES.parts?.["xl/drawings/vmlDrawing1.vml"]

if (VML === undefined) throw new Error("fixture without a drawing")

const EXCEL_IDS = {
	...NOTES,
	parts: {
		...NOTES.parts,
		"xl/drawings/vmlDrawing1.vml": { ...VML, text: `${VML_HEAD}${excelShape(-1, 1025, 1, 1)}${excelShape(0, 1026, 4, 3)}</xml>` }
	}
}

function shapeIds(vml: string): string[] {
	return [...vml.matchAll(/(?:\sid|o:spid)="(_x0000_s\d+)"/g)].map(match => match[1] ?? "")
}

describe("note shape ids", () => {
	it("gives a new note an id no shape's o:spid holds", async () => {
		const { session, sheets } = await openAnnotated(EXCEL_IDS)

		await settle()

		const sheet = sheets[0] ?? 0

		changeOf(await session.command({ type: "note.set", sheet, row: 0, col: 0, text: "new" }))

		const ids = shapeIds(entries(await savedBytes(session))["xl/drawings/vmlDrawing1.vml"] ?? "")

		expect(ids).toHaveLength(3)
		expect(new Set(ids).size).toBe(3)
		session.close()
	})

	it("gives a duplicated sheet's shapes ids of their own", async () => {
		const { session, sheets } = await openAnnotated(EXCEL_IDS)

		await settle()
		changeOf(await session.command({ type: "sheet.duplicate", sheet: sheets[0] ?? 0 }))

		const ids = Object.entries(entries(await savedBytes(session)))
			.filter(([name]) => name.endsWith(".vml"))
			.flatMap(([, text]) => shapeIds(text))

		expect(ids).toHaveLength(4)
		expect(new Set(ids).size).toBe(4)
		session.close()
	})
})
